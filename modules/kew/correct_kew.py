"""
correct_kew.py — Hiệu chỉnh số liệu KEW khi máy đo bị lỗi cấu hình
(sai CT ratio, VT ratio, systematic offset...)

Cú pháp:
    python3 correct_kew.py <input_folder> <output_folder> \
        --A-mul 100 --V-mul 1.732 --P-mul 100 --A-offset 0.5

Corrections áp dụng:   value_corrected = value_raw * multiply + offset
"""
import os
import sys
import shutil
import re
import struct
import argparse


# ─── Helpers ─────────────────────────────────────────────────────────────────

def _fmt_preserve(raw_str: str, value: float) -> str:
    """Re-format a float, preserving scientific notation if the original used it."""
    if 'E' in raw_str.upper():
        return '{:+.3E}'.format(value)
    # Try to match original decimal places
    if '.' in raw_str:
        decimals = len(raw_str.rstrip().rstrip('0').rstrip().split('.')[-1])
        decimals = max(3, min(decimals, 6))
        return f'{value:.{decimals}f}'
    return f'{value:.4f}'


def _channel_multiplier(col_name: str, corrections: dict) -> tuple[float, float]:
    """
    Return (multiply, offset) for a given INPS column name.
    corrections keys: 'A', 'V', 'P', 'Q', 'S', 'PF', or specific e.g. 'A1', 'V2'
    """
    mul, off = 1.0, 0.0
    # Map column prefix → correction group
    mapping = [
        ('AVG_A', 'A'), ('MIN_A', 'A'), ('MAX_A', 'A'), ('THDAR', 'A'),
        ('AVG_V', 'V'), ('MIN_V', 'V'), ('MAX_V', 'V'), ('THDVR', 'V'),
        ('AVG_P', 'P'), ('MIN_P', 'P'), ('MAX_P', 'P'),
        ('AVG_Q', 'Q'), ('MIN_Q', 'Q'), ('MAX_Q', 'Q'),
        ('AVG_S', 'S'), ('MIN_S', 'S'), ('MAX_S', 'S'),
        ('AVG_PF', 'PF'),
    ]
    for prefix, group in mapping:
        if col_name.startswith(prefix):
            # Check for phase-specific override first (e.g., 'A1')
            phase_char = ''
            for ph in ('1', '2', '3', 'T', 'N'):
                if col_name.startswith(prefix + ph) or col_name.startswith(prefix[4:] + ph):
                    phase_char = ph
                    break
            specific_key = group + phase_char
            if specific_key in corrections:
                c = corrections[specific_key]
            elif group in corrections:
                c = corrections[group]
            else:
                break
            mul = c.get('multiply', 1.0)
            off = c.get('offset', 0.0)
            break
    return mul, off


# ─── INPS text correction ─────────────────────────────────────────────────────

def process_inps(in_path: str, out_path: str, corrections: dict):
    try:
        with open(in_path, 'r', encoding='ascii', errors='ignore') as f:
            lines = f.readlines()

        cols = lines[1].strip('\r\n').split(',')
        out_lines = [lines[0], lines[1]]

        # Pre-compute (mul, offset) for each column index
        col_corr = [_channel_multiplier(c, corrections) for c in cols]

        for row in lines[2:]:
            parts = row.strip('\r\n').split(',')
            if len(parts) < len(cols):
                out_lines.append(row)
                continue
            for i, (mul, off) in enumerate(col_corr):
                if i >= len(parts): break
                if mul == 1.0 and off == 0.0: continue
                try:
                    v = float(parts[i])
                    corrected = v * mul + off
                    parts[i] = _fmt_preserve(parts[i], corrected)
                except ValueError:
                    pass
            out_lines.append(','.join(parts) + '\n')

        with open(out_path, 'w', encoding='ascii') as f:
            f.writelines(out_lines)
        print(f'[OK] Corrected INPS → {out_path}')
    except Exception as e:
        print(f'[ERROR] INPS correction failed: {e}')
        shutil.copy2(in_path, out_path)


# ─── INHS binary correction ───────────────────────────────────────────────────

# Maps INHS channel name → correction group
_INHS_CH_GROUP = {
    'A1[A]': 'A', 'A2[A]': 'A', 'A3[A]': 'A', 'AN[A]': 'A',
    'V1[V]': 'V', 'V2[V]': 'V', 'V3[V]': 'V',
    'P1[W]': 'P', 'P2[W]': 'P', 'P3[W]': 'P', 'P[W]': 'P',
    'Q1[var]': 'Q', 'Q2[var]': 'Q', 'Q3[var]': 'Q',
    'S1[VA]': 'S', 'S2[VA]': 'S', 'S3[VA]': 'S',
}


def _get_inhs_corr(ch_name: str, corrections: dict) -> tuple[float, float]:
    group = _INHS_CH_GROUP.get(ch_name)
    if not group:
        return 1.0, 0.0
    # Phase-specific key (A1, V2, etc.)
    phase_char = ch_name[1] if len(ch_name) > 1 and ch_name[1].isdigit() else ''
    specific = group + phase_char
    c = corrections.get(specific, corrections.get(group, {}))
    return c.get('multiply', 1.0), c.get('offset', 0.0)


def _apply_inhs_correction(payload: bytes, r_parts: list, mul: float, off: float) -> bytes:
    try:
        size = int(r_parts[5].decode('ascii').strip())
        p_strip = payload[:-2] if payload.endswith(b'\r\n') else payload
        num_vals = len(p_strip) // size
        fmt = f'<{num_vals}i' if size == 4 else f'<{num_vals}h'
        data_bytes = p_strip[:num_vals * size]
        vals = struct.unpack(fmt, data_bytes)
        corrected = [int(v * mul + off) for v in vals]
        packed = struct.pack(fmt, *corrected)
        return packed + payload[num_vals * size:]
    except Exception:
        return payload


def process_inhs(in_path: str, out_path: str, corrections: dict):
    try:
        with open(in_path, 'rb') as f:
            content = f.read()

        lines = content.split(b'\r\n', 2)
        if len(lines) < 3:
            shutil.copy2(in_path, out_path)
            return

        magic, header, data_section = lines[0], lines[1], lines[2]
        record_positions = [m.start() for m in re.finditer(rb'20\d\d/\d\d/\d\d', data_section)]
        if not record_positions:
            shutil.copy2(in_path, out_path)
            return

        out_records = []
        for i, start in enumerate(record_positions):
            end = record_positions[i + 1] if i + 1 < len(record_positions) else len(data_section)
            rec = data_section[start:end]
            parts = rec.split(b',', 6)
            if len(parts) < 7:
                out_records.append(rec)
                continue
            ch_str = parts[3].decode('ascii', errors='ignore').strip()
            mul, off = _get_inhs_corr(ch_str, corrections)
            if mul == 1.0 and off == 0.0:
                out_records.append(rec)
            else:
                new_payload = _apply_inhs_correction(parts[6], parts, mul, off)
                out_records.append(b','.join(parts[:6]) + b',' + new_payload)

        with open(out_path, 'wb') as f:
            f.write(magic + b'\r\n' + header + b'\r\n' + b''.join(out_records))
        print(f'[OK] Corrected INHS → {out_path}')
    except Exception as e:
        print(f'[ERROR] INHS correction failed: {e}')
        shutil.copy2(in_path, out_path)


# ─── Folder-level correction ──────────────────────────────────────────────────

def process_folder(in_folder: str, out_folder: str, corrections: dict):
    """Apply corrections to all KEW files in a folder and write to out_folder."""
    os.makedirs(out_folder, exist_ok=True)
    for fname in os.listdir(in_folder):
        if not fname.upper().endswith('.KEW'):
            continue
        in_p = os.path.join(in_folder, fname)
        out_p = os.path.join(out_folder, fname)
        upper = fname.upper()
        if upper.startswith('INPS'):
            process_inps(in_p, out_p, corrections)
        elif upper.startswith('INHS'):
            process_inhs(in_p, out_p, corrections)
        else:
            shutil.copy2(in_p, out_p)
            print(f'[OK] Copied {fname}')


# ─── CLI ─────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Hiệu chỉnh số liệu KEW')
    parser.add_argument('input_folder')
    parser.add_argument('output_folder')
    # Multiply flags
    for grp in ('A', 'V', 'P', 'Q', 'S'):
        parser.add_argument(f'--{grp}-mul', type=float, default=1.0, help=f'Nhân hệ số cho kênh {grp}')
        parser.add_argument(f'--{grp}-offset', type=float, default=0.0, help=f'Cộng offset cho kênh {grp}')
    args = parser.parse_args()

    corrections = {}
    for grp in ('A', 'V', 'P', 'Q', 'S'):
        mul = getattr(args, f'{grp}_mul')
        off = getattr(args, f'{grp}_offset')
        if mul != 1.0 or off != 0.0:
            corrections[grp] = {'multiply': mul, 'offset': off}

    if not corrections:
        print('Không có hiệu chỉnh nào được chỉ định. Dùng --A-mul, --V-mul, ...')
        sys.exit(1)

    print(f'Bắt đầu hiệu chỉnh: {corrections}')
    process_folder(args.input_folder, args.output_folder, corrections)
    print('Hoàn tất.')
