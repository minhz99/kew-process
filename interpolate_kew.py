import os
import sys
import shutil
import re
import struct
import random
import math

# ─── Ornstein-Uhlenbeck Noise Generator ──────────────────────────────────────
# Generates correlated random-walk noise that mimics real electrical measurement
# drift. Unlike pure white noise, OU process has a "memory" so values drift
# slowly and revert to the mean — exactly like real-world phase imbalance.
class OUProcess:
    """
    Ornstein-Uhlenbeck process: dx = theta*(mu-x)*dt + sigma*dW
    - theta: mean reversion speed (higher = faster revert to mu)
    - mu:    long-run mean (1.0 = no bias)
    - sigma: volatility
    """
    def __init__(self, theta=0.3, mu=1.0, sigma=0.015, dt=1.0):
        self.theta = theta
        self.mu = mu
        self.sigma = sigma
        self.dt = dt
        self.x = mu + random.gauss(0, sigma)

    def step(self):
        """Advance by one step and return the current multiplier."""
        dW = random.gauss(0, math.sqrt(self.dt))
        self.x += self.theta * (self.mu - self.x) * self.dt + self.sigma * dW
        # Clamp to [0.90, 1.10] so we never create physically impossible values
        self.x = max(0.90, min(1.10, self.x))
        return self.x


def get_ref_phase_inhs(data_section, record_positions):
    """Xác định pha có dòng điện hoạt động mạnh nhất bằng cách đo kích thước byte khác 0"""
    sums = {'1': 0, '2': 0, '3': 0}
    for i in range(len(record_positions)):
        start = record_positions[i]
        end = record_positions[i+1] if i+1 < len(record_positions) else len(data_section)
        rec = data_section[start:end]
        parts = rec.split(b',', 6)
        if len(parts) > 6:
            ch = parts[3].decode('ascii', errors='ignore').strip()
            if ch in ('A1[A]', 'A2[A]', 'A3[A]'):
                phase = ch[1]
                payload = parts[6]
                sums[phase] += sum(1 for b in payload if b != 0 and b != 13 and b != 10)
    
    ref = max(sums, key=sums.get)
    print(f"[INFO] INHS Reference phase detected: {ref}")
    return ref


def process_inhs(in_path, out_path):
    with open(in_path, 'rb') as f:
        content = f.read()
    
    lines = content.split(b'\r\n', 2)
    if len(lines) < 3:
        shutil.copy2(in_path, out_path)
        return None
        
    magic = lines[0]
    header = lines[1]
    data_section = lines[2]
    
    record_positions = [m.start() for m in re.finditer(rb'20\d\d/\d\d/\d\d', data_section)]
    if not record_positions:
        shutil.copy2(in_path, out_path)
        return None
        
    ref_phase = get_ref_phase_inhs(data_section, record_positions)
    
    ref_A_ch = f'A{ref_phase}[A]'.encode('ascii')
    ref_P_ch = f'P{ref_phase}[W]'.encode('ascii')
    
    # Build per-phase OU processes so noise is correlated across time
    ou_A2 = OUProcess(theta=0.25, sigma=0.012)
    ou_A3 = OUProcess(theta=0.20, sigma=0.010)
    ou_P2 = OUProcess(theta=0.25, sigma=0.012)
    ou_P3 = OUProcess(theta=0.20, sigma=0.010)
    
    # Parse records
    records = []
    for i in range(len(record_positions)):
        start = record_positions[i]
        end = record_positions[i+1] if i+1 < len(record_positions) else len(data_section)
        rec = data_section[start:end]
        parts = rec.split(b',', 6)
        if len(parts) > 6:
            dt = parts[0] + b' ' + parts[1]
            ch = parts[3].strip()
            records.append({'dt': dt, 'ch': ch, 'parts': parts, 'raw': rec})
        else:
            records.append({'dt': None, 'raw': rec})
            
    out_records = []
    current_dt = None
    group = []
    
    def apply_ou_noise(payload, r_parts, ou_proc):
        """Apply OU-correlated noise to a binary payload."""
        try:
            size = int(r_parts[5].decode('ascii').strip())
            p_strip = payload[:-2] if payload.endswith(b'\r\n') else payload
            num_vals = len(p_strip) // size
            fmt = f'<{num_vals}i' if size == 4 else f'<{num_vals}h'
            
            data_bytes = p_strip[:num_vals*size]
            vals = struct.unpack(fmt, data_bytes)
            
            # One OU step per record group — same multiplier for all samples in this waveform
            multiplier = ou_proc.step()
            noisy_vals = [int(v * multiplier) if abs(v) > 0 else v for v in vals]
            packed = struct.pack(fmt, *noisy_vals)
            return packed + payload[num_vals*size:]
        except Exception:
            return payload

    def process_group(g):
        payload_A = None
        payload_P = None
        payload_A_parts = None
        payload_P_parts = None
        for r in g:
            if r.get('ch') == ref_A_ch:
                payload_A = r['parts'][6]
                payload_A_parts = r['parts']
            elif r.get('ch') == ref_P_ch:
                payload_P = r['parts'][6]
                payload_P_parts = r['parts']
                
        for r in g:
            if not r.get('dt'):
                out_records.append(r['raw'])
                continue
                
            ch_str = r['ch'].decode('ascii').strip()
            new_payload = None
            
            if ch_str in ('A1[A]', 'A2[A]', 'A3[A]') and payload_A:
                if ch_str == f'A{ref_phase}[A]':
                    new_payload = payload_A
                elif ch_str == 'A2[A]':
                    new_payload = apply_ou_noise(payload_A, payload_A_parts, ou_A2)
                else:
                    new_payload = apply_ou_noise(payload_A, payload_A_parts, ou_A3)
            elif ch_str in ('P1[W]', 'P2[W]', 'P3[W]') and payload_P:
                if ch_str == f'P{ref_phase}[W]':
                    new_payload = payload_P
                elif ch_str == 'P2[W]':
                    new_payload = apply_ou_noise(payload_P, payload_P_parts, ou_P2)
                else:
                    new_payload = apply_ou_noise(payload_P, payload_P_parts, ou_P3)
                
            if new_payload:
                new_raw = b','.join(r['parts'][:6]) + b',' + new_payload
                out_records.append(new_raw)
            else:
                out_records.append(r['raw'])

    for r in records:
        if r['dt'] != current_dt:
            if group:
                process_group(group)
            group = [r]
            current_dt = r['dt']
        else:
            group.append(r)
    if group:
        process_group(group)
        
    with open(out_path, 'wb') as f:
        f.write(magic + b'\r\n' + header + b'\r\n' + b''.join(out_records))
    print(f"[OK] Processed INHS: {out_path}")
    return ref_phase


def process_inps(in_path, out_path, ref_phase=None):
    try:
        with open(in_path, 'r', encoding='ascii', errors='ignore') as f:
            lines = f.readlines()
            
        header_line = lines[1].strip('\r\n')
        cols = header_line.split(',')
        
        if ref_phase is None:
            a1_idx = cols.index('AVG_A1[A]') if 'AVG_A1[A]' in cols else -1
            a2_idx = cols.index('AVG_A2[A]') if 'AVG_A2[A]' in cols else -1
            a3_idx = cols.index('AVG_A3[A]') if 'AVG_A3[A]' in cols else -1
            
            sums = {'1': 0.0, '2': 0.0, '3': 0.0}
            for row in lines[2:]:
                r_parts = row.strip('\r\n').split(',')
                for phase, idx in [('1', a1_idx), ('2', a2_idx), ('3', a3_idx)]:
                    if idx != -1 and idx < len(r_parts):
                        try: sums[phase] += float(r_parts[idx])
                        except: pass
                        
            ref_phase = max(sums, key=sums.get)
            print(f"[INFO] INPS Reference phase detected: {ref_phase}")
        else:
            print(f"[INFO] INPS using INHS Reference phase: {ref_phase}")

        # Build separate OU processes per target phase, for each metric group
        ou_ph2 = OUProcess(theta=0.3, sigma=0.008)
        ou_ph3 = OUProcess(theta=0.25, sigma=0.007)
        
        out_lines = [lines[0], lines[1]]
        
        prefixes_to_copy = ['AVG_A', 'AVG_P', 'AVG_Q', 'AVG_S', 'AVG_PF', 'THDAR']
        
        copy_map = {}  # tgt_idx -> (src_idx, target_phase)
        
        for g in prefixes_to_copy:
            for src_idx, c in enumerate(cols):
                if c.startswith(g) and f"{g}{ref_phase}" in c:
                    basename = c.replace(f"{g}{ref_phase}", "")
                    for offset in range(3):
                        s_i = src_idx + offset
                        if s_i >= len(cols): break
                        for p in ('1', '2', '3'):
                            if p == ref_phase: continue
                            tgt_col = f"{g}{p}{basename}"
                            try:
                                t_i_col = cols.index(tgt_col)
                                t_i = t_i_col + offset
                                if t_i < len(cols):
                                    copy_map[t_i] = (s_i, p)
                            except ValueError:
                                pass
                                
        totals_map = {}
        for g in ['AVG_P', 'AVG_Q', 'AVG_S']:
            ref_col = next((c for c in cols if c.startswith(f"{g}{ref_phase}")), None)
            tot_col = next((c for c in cols if c.startswith(f"{g}[")), None)
            
            if ref_col and tot_col:
                r_idx = cols.index(ref_col)
                t_idx = cols.index(tot_col)
                for offset in range(3):
                    if t_idx + offset < len(cols) and r_idx + offset < len(cols):
                        totals_map[t_idx+offset] = r_idx+offset
        
        for row in lines[2:]:
            r_parts = row.strip('\r\n').split(',')
            if not r_parts or len(r_parts) < len(cols):
                out_lines.append(row)
                continue
            
            # Advance OU per row (one step per 1-second interval)
            m2 = ou_ph2.step()
            m3 = ou_ph3.step()
            
            for t_idx, (s_idx, tgt_phase) in copy_map.items():
                if t_idx < len(r_parts) and s_idx < len(r_parts):
                    try:
                        v = float(r_parts[s_idx])
                        multiplier = m2 if tgt_phase == '2' else m3
                        if v != 0:
                            v *= multiplier
                        fmt_sci = 'E' in r_parts[s_idx].upper()
                        r_parts[t_idx] = '{:+.3E}'.format(v) if fmt_sci else f"{v:.4f}"
                    except ValueError:
                        r_parts[t_idx] = r_parts[s_idx]
                    
            # Totals must reflect the sum of the three (different) phases, not 3x ref
            for t_idx, s_idx in totals_map.items():
                if t_idx < len(r_parts) and s_idx < len(r_parts):
                    try:
                        v_ref = float(r_parts[s_idx])
                        # Recover phase 2 and 3 values we just wrote
                        # Hack: just compute 3-phase total based on individual
                        # values already set in r_parts for the two copies
                        # Find the three per-phase source indices for this total
                        v_tot = v_ref * (1 + m2 + m3)  # ref + phase2 multiplied + phase3 multiplied
                        fmt_sci = 'E' in r_parts[s_idx].upper()
                        r_parts[t_idx] = '{:+.3E}'.format(v_tot) if fmt_sci else f"{v_tot:.4f}"
                    except ValueError:
                        pass
                        
            out_lines.append(','.join(r_parts) + '\n')
            
        with open(out_path, 'w', encoding='ascii') as f:
            f.writelines(out_lines)
            
        print(f"[OK] Processed INPS: {out_path}")
        
    except Exception as e:
        print(f"[ERROR] Process INPS failed: {e}")
        import traceback; traceback.print_exc()
        shutil.copy2(in_path, out_path)


def detect_missing_phases(folder):
    """Returns dict: {'ref_phase': '3', 'missing': ['1', '2'], 'has_inhs': True}"""
    inhs_file = next((f for f in os.listdir(folder) if f.upper().startswith('INHS') and f.upper().endswith('.KEW')), None)
    if not inhs_file:
        return {'ref_phase': None, 'missing': [], 'has_inhs': False}
    
    path = os.path.join(folder, inhs_file)
    with open(path, 'rb') as f:
        content = f.read()
    
    lines = content.split(b'\r\n', 2)
    if len(lines) < 3:
        return {'ref_phase': None, 'missing': [], 'has_inhs': True}
    
    data_section = lines[2]
    record_positions = [m.start() for m in re.finditer(rb'20\d\d/\d\d/\d\d', data_section)]
    if not record_positions:
        return {'ref_phase': None, 'missing': [], 'has_inhs': True}
    
    sums = {'1': 0, '2': 0, '3': 0}
    for i in range(len(record_positions)):
        start = record_positions[i]
        end = record_positions[i+1] if i+1 < len(record_positions) else len(data_section)
        rec = data_section[start:end]
        parts = rec.split(b',', 6)
        if len(parts) > 6:
            ch = parts[3].decode('ascii', errors='ignore').strip()
            if ch in ('A1[A]', 'A2[A]', 'A3[A]'):
                phase = ch[1]
                payload = parts[6]
                sums[phase] += sum(1 for b in payload if b != 0 and b != 13 and b != 10)
    
    total = sum(sums.values())
    if total == 0:
        return {'ref_phase': None, 'missing': [], 'has_inhs': True}
    
    ref_phase = max(sums, key=sums.get)
    # A phase is "missing" if it has <5% of the reference phase activity
    ref_activity = sums[ref_phase]
    missing = [p for p in ('1', '2', '3') if p != ref_phase and sums[p] < ref_activity * 0.05]
    
    return {
        'ref_phase': ref_phase,
        'missing': missing,
        'has_inhs': True,
        'phase_activity': sums
    }


def process_folder(in_folder, out_folder):
    if not os.path.exists(out_folder):
        os.makedirs(out_folder)
        
    in_files = os.listdir(in_folder)
    
    inhs_file = next((f for f in in_files if f.upper().startswith('INHS') and f.upper().endswith('.KEW')), None)
    inps_file = next((f for f in in_files if f.upper().startswith('INPS') and f.upper().endswith('.KEW')), None)
    
    ref_phase = None
    if inhs_file:
        in_p = os.path.join(in_folder, inhs_file)
        out_p = os.path.join(out_folder, inhs_file)
        ref_phase = process_inhs(in_p, out_p)
        
    if inps_file:
        in_p = os.path.join(in_folder, inps_file)
        out_p = os.path.join(out_folder, inps_file)
        process_inps(in_p, out_p, ref_phase=ref_phase)
        
    for f in in_files:
        if not f.upper().endswith('.KEW'):
            continue
        if f == inhs_file or f == inps_file:
            continue
        in_p = os.path.join(in_folder, f)
        out_p = os.path.join(out_folder, f)
        shutil.copy2(in_p, out_p)
        print(f"[OK] Copied {f}")


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Usage: python3 interpolate_kew.py <input_folder> <output_folder>")
        sys.exit(1)
    print(f"Bắt đầu nội suy từ {sys.argv[1]} sang {sys.argv[2]}")
    process_folder(sys.argv[1], sys.argv[2])
    print("Hoàn tất.")
