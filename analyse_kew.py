"""
analyse_kew.py - Phân tích dữ liệu power quality từ KEW files
Xuất ra file JSON để dùng trong dashboard HTML
Hỗ trợ đầy đủ bộ file: INHS, INPS, EVTS, INIS, SUPS
"""

import struct
import pandas as pd
import numpy as np
import json
import os
import sys
import re

# ─────────────────────────────────────────────────────────
# Bước 1: Đọc INHS file (binary mixed format) - hàm từ read-kew.py
# ─────────────────────────────────────────────────────────

def parse_kew_file(filepath):
    """
    Parse KEW file - hỗ trợ cả text và binary mixed format (INHS, EVTS...).
    Sử dụng Robust Regex Splitting để tách bản ghi tránh lỗi do byte đặc biệt trong data binary.
    """
    try:
        with open(filepath, 'rb') as f:
            content = f.read()
        if not content:
            return None, None
            
        # Tách magic header và column headers (thường ở 2 dòng đầu)
        lines = content.split(b'\r\n', 2)
        if len(lines) < 2:
            # Thử split với \n đơn thuần
            lines = content.split(b'\n', 2)
            
        magic = lines[0].decode('ascii', errors='ignore').strip()
        if magic.startswith('FILE ID'):
            return magic, None
            
        header_bytes = lines[1].strip()
        header_str = header_bytes.decode('ascii', errors='ignore')
        if header_str.endswith(','):
            header_str = header_str[:-1]
        headers = [h.strip() for h in header_str.split(',')]
        
        has_k = 'k' in headers
        has_size = 'SIZE' in headers
        text_cols_count = (headers.index('SIZE') + 1) if has_size else len(headers)
        
        # Phần dữ liệu thô sau Header
        data_section = lines[2] if len(lines) > 2 else b''
        
        # Sử dụng Regex để tìm vị trí bắt đầu mỗi bản ghi (định dạng ngày 20XX/XX/XX)
        # Cách này cực kỳ bền bỉ vì không phụ thuộc vào ký tự xuống dòng (CRLF) trong vùng binary.
        record_positions = [m.start() for m in re.finditer(rb'20\d\d/\d\d/\d\d', data_section)]
        
        data_rows = []
        for i, start in enumerate(record_positions):
            end = record_positions[i+1] if i+1 < len(record_positions) else len(data_section)
            record_raw = data_section[start:end].strip(b'\r\n')
            if not record_raw:
                continue
                
            if has_k and has_size:
                # Binary mixed format (ví dụ INHS)
                parts = record_raw.split(b',', text_cols_count)
                if len(parts) <= text_cols_count:
                    continue
                
                text_parts = [p.decode('ascii', errors='ignore').strip() for p in parts[:text_cols_count]]
                try:
                    k = float(text_parts[headers.index('k')])
                    size = int(text_parts[headers.index('SIZE')])
                except (ValueError, IndexError):
                    continue
                
                binary_data = parts[text_cols_count]
                num_bin_values = len(headers) - text_cols_count
                row_data = text_parts.copy()
                
                expected_bytes = num_bin_values * size
                if len(binary_data) >= expected_bytes:
                    fmt = f'<{num_bin_values}i' if size == 4 else f'<{num_bin_values}h'
                    vals = struct.unpack(fmt, binary_data[:expected_bytes])
                    row_data.extend([v * k for v in vals])
                else:
                    row_data.extend([None] * num_bin_values)
                data_rows.append(row_data)
            else:
                # Text only format (ví dụ EVTS)
                line_str = record_raw.decode('ascii', errors='ignore')
                if line_str.endswith(','):
                    line_str = line_str[:-1]
                data_rows.append(line_str.split(','))
        
        df = pd.DataFrame(data_rows, columns=headers)
        df.columns = [c.strip() for c in df.columns]
        
        if 'DATE' in df.columns and 'TIME' in df.columns:
            df['DATETIME'] = pd.to_datetime(
                df['DATE'].astype(str) + ' ' + df['TIME'].astype(str),
                format='%Y/%m/%d %H:%M:%S.%f', errors='coerce'
            )
            mask = df['DATETIME'].isna()
            if mask.any():
                df.loc[mask, 'DATETIME'] = pd.to_datetime(
                    df.loc[mask, 'DATE'].astype(str) + ' ' + df.loc[mask, 'TIME'].astype(str),
                    format='%Y/%m/%d %H:%M:%S', errors='coerce'
                )
        
        for col in df.columns:
            if col not in ['DATE', 'TIME', 'DATETIME', 'ELAPSED TIME', 'CH']:
                if df[col].dtype == object:
                    converted = pd.to_numeric(df[col], errors='coerce')
                    if not converted.isna().all():
                        df[col] = converted
        
        return magic, df
    except Exception as e:
        print(f"[ERROR] parse_kew_file failed: {e}")
        return None, None



# ─────────────────────────────────────────────────────────
# Bước 2: Phân tích INHS data
# ─────────────────────────────────────────────────────────

def analyse_inhs(df):
    """Phân tích harmonics + fundamental values theo thời gian"""
    
    harm_cols = [f'AVG_{i:02d}' for i in range(1, 51)]
    results = {}
    
    channels_of_interest = ['V1[V]', 'V2[V]', 'V3[V]', 'A1[A]', 'A2[A]', 'A3[A]',
                             'P1[W]', 'P2[W]', 'P3[W]', 'P[W]',
                             'VA1[deg]', 'VA2[deg]', 'VA3[deg]']
    
    for ch in df['CH'].unique():
        ch_clean = ch.strip()
        ch_df = df[df['CH'] == ch].sort_values('DATETIME')
        
        # Fundamental (1st harmonic = AVG_01) theo thời gian
        # Giữ lại NaNs để charts có thể hiển thị khoảng trống (gaps)
        fundamental = ch_df['AVG_01'].values.astype(float)
        timestamps = ch_df['DATETIME'].dt.strftime('%Y-%m-%d %H:%M:%S').tolist()
        
        # THD = sqrt(sum(H2..H50)^2) / H1 * 100
        # Không fillna(0) toàn bộ để tránh làm sai lệch biểu đồ ở những vùng mất dữ liệu
        harmonics = ch_df[harm_cols].values.astype(float)
        h1 = harmonics[:, 0]
        h_rest = harmonics[:, 1:]
        
        with np.errstate(divide='ignore', invalid='ignore'):
            sum_h_rest_sq = np.nansum(h_rest**2, axis=1)
            # Chỉ tính THD khi có fundamental h1 và h1 > 0
            # Những vùng h1 là NaN hoặc 0 sẽ giữ là NaN trong thd array
            thd = np.where((~np.isnan(h1)) & (h1 > 0), 
                           np.sqrt(sum_h_rest_sq) / h1 * 100, 
                           np.nan)
        
        # Harmonic spectrum – average over recording
        avg_spectrum = np.nanmean(harmonics, axis=0).tolist()
        
        results[ch_clean] = {
            'timestamps': timestamps,
            'fundamental': fundamental.tolist(),
            'thd': thd.tolist(),
            'avg_thd': float(np.nanmean(thd)) if not np.all(np.isnan(thd)) else None,
            'avg_fundamental': float(np.nanmean(h1)) if not np.all(np.isnan(h1)) else None,
            'min_fundamental': float(np.nanmin(h1)) if not np.all(np.isnan(h1)) else None,
            'max_fundamental': float(np.nanmax(h1)) if not np.all(np.isnan(h1)) else None,
            'spectrum': avg_spectrum,  # 50 harmonics averages
        }
    
    return results


# ─────────────────────────────────────────────────────────
# Bước 3: Đọc EVTS file
# ─────────────────────────────────────────────────────────

def parse_evts(filepath):
    try:
        _, df = parse_kew_file(filepath)
        if df is None:
            return []
        df.columns = [c.strip() for c in df.columns]
        
        events = []
        for _, row in df.iterrows():
            ev = {
                'datetime': str(row.get('DATETIME', '')),
                'elapsed': str(row.get('ELAPSED TIME', '')),
            }
            
            event_types = {
                'Transient[S/E]': ('Transient', 'Transient[V]'),
                'Intrupt[S/E]': ('Interrupt', 'Intrupt[V]'),
                'Dip[S/E]': ('Voltage Dip', 'Dip[V]'),
                'Swell[S/E]': ('Voltage Swell', 'Swell[V]'),
                'Inrush current[S/E]': ('Inrush Current', 'Inrush current[A]'),
            }
            
            for col, (name, val_col) in event_types.items():
                if col in row:
                    val_se = row[col]
                    # Could be string '1' or int 1
                    is_active = (str(val_se).strip() == '1')
                    if is_active:
                        val = row.get(val_col, '')
                        ev_entry = dict(ev)
                        ev_entry['type'] = name
                        try:
                            ev_entry['value'] = float(str(val).strip())
                        except (ValueError, TypeError):
                            ev_entry['value'] = None
                        events.append(ev_entry)
        
        return events
    except Exception as e:
        print(f"[WARN] Could not parse EVTS: {e}")
        return []


# ─────────────────────────────────────────────────────────
# Bước 4: Đọc INIS (metadata)
# ─────────────────────────────────────────────────────────

def parse_inis(filepath):
    meta = {}
    try:
        with open(filepath, 'r', encoding='ascii', errors='ignore') as f:
            for line in f:
                line = line.strip()
                if ',:,' in line:
                    parts = line.split(',:,')
                    if len(parts) == 2:
                        key = parts[0].strip()
                        val = parts[1].strip().strip("'").rstrip(',')
                        meta[key] = val
    except Exception:
        pass
    return meta


# ─────────────────────────────────────────────────────────
# Bước 4b: Đọc SUPS (cấu hình thiết bị, INI format)
# ─────────────────────────────────────────────────────────

def parse_sups(filepath):
    """Đọc file SUPS - cấu hình thiết bị dạng INI"""
    config = {}
    current_section = 'General'
    try:
        with open(filepath, 'r', encoding='ascii', errors='ignore') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                if line.startswith('[') and line.endswith(']'):
                    current_section = line[1:-1]
                    config.setdefault(current_section, {})
                elif '=' in line:
                    key, _, val = line.partition('=')
                    config.setdefault(current_section, {})[key.strip()] = val.strip()
    except Exception:
        pass
    return config


# ─────────────────────────────────────────────────────────
# Bước 4c: Đọc INPS (Interval Power Series - CSV text)
# Format: DATE,TIME,ELAPSED TIME,AVG_V1[V],,,AVG_V2[V],,,..  
# Mỗi đại lượng gồm 3 giá trị: giá trị (avg), min, max (ba cột liền)
# ─────────────────────────────────────────────────────────

def parse_inps(filepath):
    """Đọc INPS file - dữ liệu power quality trung bình theo thời gian (1s/mẫu)"""
    try:
        with open(filepath, 'rb') as f:
            magic_bytes = f.readline().strip()
            magic = magic_bytes.decode('ascii', errors='ignore')
            
            header_bytes = f.readline()
            header_str = header_bytes.decode('ascii', errors='ignore').strip()
            
            # Tách header: cột dạng "AVG_V1[V],,," = 3 slots (avg, min, max)
            # Xây dựng danh sách tên cột đầy đủ
            raw_cols = header_str.split(',')
            expanded_cols = []
            i = 0
            while i < len(raw_cols):
                col = raw_cols[i].strip()
                if col == '':
                    # Cột trống sau tên đại lượng = min, max
                    if expanded_cols:
                        base = expanded_cols[-1].split('_avg')[0] if '_avg' in expanded_cols[-1] else expanded_cols[-1]
                        count_empty = 0
                        j = i
                        while j < len(raw_cols) and raw_cols[j].strip() == '':
                            count_empty += 1
                            j += 1
                        if count_empty >= 2:
                            expanded_cols.append(base + '_min')
                            expanded_cols.append(base + '_max')
                            i += 2
                            continue
                        else:
                            expanded_cols.append('')
                    else:
                        expanded_cols.append('')
                else:
                    expanded_cols.append(col)
                i += 1
            
            # Đọc dữ liệu
            data_rows = []
            for line in f:
                line_str = line.decode('ascii', errors='ignore').strip()
                if not line_str:
                    continue
                parts = line_str.split(',')
                data_rows.append(parts)
            
            # Đặt tên duy nhất cho cột rỗng để tránh lỗi DataFrame
            final_cols = []
            blank_count = 0
            for col in expanded_cols[:len(raw_cols)]:
                if col == '':
                    blank_count += 1
                    final_cols.append(f'_blank_{blank_count}')
                else:
                    final_cols.append(col)
            
            # Tạo DataFrame
            n_cols = len(raw_cols)
            padded_rows = []
            for row in data_rows:
                if len(row) < n_cols:
                    row = row + [''] * (n_cols - len(row))
                padded_rows.append(row[:n_cols])
            
            df = pd.DataFrame(padded_rows, columns=final_cols)
            
            # Parse datetime
            if 'DATE' in df.columns and 'TIME' in df.columns:
                df['DATETIME'] = pd.to_datetime(
                    df['DATE'].astype(str) + ' ' + df['TIME'].astype(str),
                    format='%Y/%m/%d %H:%M:%S', errors='coerce'
                )
            
            # Parse numeric (bỏ qua cột placeholder _blank_)
            skip = {'DATE', 'TIME', 'DATETIME', 'ELAPSED TIME'}
            for col in df.columns:
                if col in skip or col.startswith('_blank_'):
                    continue
                if df[col].dtype == object:
                    converted = pd.to_numeric(df[col], errors='coerce')
                    if not converted.isna().all():
                        df[col] = converted
            
            return magic, df
    
    except FileNotFoundError:
        print(f"[ERROR] INPS file not found: {filepath}")
        return None, None
    except Exception as e:
        print(f"[WARN] Could not parse INPS: {e}")
        return None, None


def analyse_inps(df):
    """Tổng hợp thống kê từ INPS dataframe"""
    if df is None or df.empty:
        return {}
    
    result = {}
    skip_cols = {'DATE', 'TIME', 'DATETIME', 'ELAPSED TIME', 'ELAPSED TIME_min', 'ELAPSED TIME_max'}
    
    # Lấy các cột avg (không phải _min/_max)
    avg_cols = [c for c in df.columns 
                if c not in skip_cols 
                and not c.endswith('_min') 
                and not c.endswith('_max')
                and c.strip() != '']
    
    timestamps = []
    if 'DATETIME' in df.columns:
        timestamps = df['DATETIME'].dt.strftime('%Y-%m-%d %H:%M:%S').tolist()
    
    for col in avg_cols:
        try:
            vals = pd.to_numeric(df[col], errors='coerce')
            min_col = col + '_min'
            max_col = col + '_max'
            
            entry = {
                'timestamps': timestamps,
                'values': vals.tolist(),
                'avg': float(vals.mean()) if not vals.isna().all() else None,
                'min': float(vals.min()) if not vals.isna().all() else None,
                'max': float(vals.max()) if not vals.isna().all() else None,
            }
            
            if min_col in df.columns:
                min_vals = pd.to_numeric(df[min_col], errors='coerce')
                entry['recorded_min'] = float(min_vals.min()) if not min_vals.isna().all() else None
            if max_col in df.columns:
                max_vals = pd.to_numeric(df[max_col], errors='coerce')
                entry['recorded_max'] = float(max_vals.max()) if not max_vals.isna().all() else None
            
            result[col] = entry
        except Exception:
            continue
    
    return result


# ─────────────────────────────────────────────────────────
# Bước 4d: Đọc VALS (RMS waveform tại thời điểm sự kiện)
# Format: mỗi record = DATE,TIME,ELAPSED,CH,scale,size,<4B header><N×int16>
# KHÔNG dùng readline() vì binary blob có thể chứa byte \n
# ─────────────────────────────────────────────────────────

def _parse_binary_event_file(filepath, label):
    """Hàm chung parse VALS và WAVS: đọc toàn bộ file, tách record bằng regex."""
    records = []
    try:
        with open(filepath, 'rb') as f:
            content = f.read()
        
        # Bỏ magic header
        first_nl = content.index(b'\r\n') + 2
        data_section = content[first_nl:]
        
        # Tìm vị trí bắt đầu mỗi record (ngày theo mẫu 20XX/XX/XX)
        record_positions = [m.start() for m in re.finditer(rb'20\d\d/\d\d/\d\d,', data_section)]
        
        for i, start in enumerate(record_positions):
            end = record_positions[i+1] if i+1 < len(record_positions) else len(data_section)
            record_raw = data_section[start:end]
            
            # Tách header text và binary blob
            parts = record_raw.split(b',', 6)
            if len(parts) < 7:
                continue
            
            try:
                date_s = parts[0].decode('ascii', errors='ignore').strip()
                time_s = parts[1].decode('ascii', errors='ignore').strip()
                elapsed_s = parts[2].decode('ascii', errors='ignore').strip()
                ch = parts[3].decode('ascii', errors='ignore').strip()
                scale = float(parts[4].decode('ascii', errors='ignore').strip())
                size = int(parts[5].decode('ascii', errors='ignore').strip())
                blob = parts[6]
                
                # VALS có 4 bytes header trước samples; WAVS không có
                # Phân biệt: VALS blob bắt đầu bằng 4 byte marker đặc biệt (thường là \x00\xCA\x9A\x3B)
                header_bytes = 0
                if label == 'VALS' and len(blob) >= 4:
                    # Skip 4-byte record header
                    header_bytes = 4
                
                samples_raw = blob[header_bytes:]
                n = len(samples_raw) // size
                
                if size == 2 and n > 0:
                    values_raw = struct.unpack(f'<{n}h', samples_raw[:n * 2])
                    values_scaled = [v * scale for v in values_raw]
                elif size == 4 and n > 0:
                    values_raw = struct.unpack(f'<{n}i', samples_raw[:n * 4])
                    values_scaled = [v * scale for v in values_raw]
                else:
                    values_scaled = []
                
                records.append({
                    'datetime': f'{date_s} {time_s}',
                    'elapsed': elapsed_s,
                    'channel': ch,
                    'scale': scale,
                    'n_samples': n,
                    'values': values_scaled,
                })
            except Exception:
                continue
        
        return records
    
    except FileNotFoundError:
        print(f'[ERROR] File not found: {filepath}')
        return []
    except Exception as e:
        print(f'[WARN] Could not parse {label}: {e}')
        return []


def parse_vals(filepath):
    """Đọc VALS - RMS values tại thời điểm sự kiện (cycle-by-cycle, ~182 samples/kênh)"""
    records = _parse_binary_event_file(filepath, 'VALS')
    print(f'[INFO] VALS parsed: {len(records)} channel records')
    return records


def parse_wavs(filepath):
    """Đọc WAVS - dạng sóng AC raw tại thời điểm sự kiện (~9604 samples/kênh, 50Hz → ~192 chu kỳ)"""
    records = _parse_binary_event_file(filepath, 'WAVS')
    print(f'[INFO] WAVS parsed: {len(records)} channel records')
    return records


# ─────────────────────────────────────────────────────────
# Bước 5: Tổng hợp + xuất JSON
# ─────────────────────────────────────────────────────────

def find_file(folder, prefix):
    for f in os.listdir(folder):
        if f.upper().startswith(prefix.upper()) and f.upper().endswith('.KEW'):
            return os.path.join(folder, f)
    return None

def build_analysis(folder):
    print(f"[INFO] Analysing folder: {folder}")
    folder_name = os.path.basename(folder.rstrip(os.sep))
    
    inhs_path = find_file(folder, 'INHS')
    evts_path = find_file(folder, 'EVTS')
    inis_path = find_file(folder, 'INIS')
    inps_path = find_file(folder, 'INPS')
    sups_path = find_file(folder, 'SUPS')
    vals_path = find_file(folder, 'VALS')
    wavs_path = find_file(folder, 'WAVS')
    
    if not inhs_path:
        print("[ERROR] No INHS data file found in the folder.")
        return None
    
    # --- Metadata ---
    meta = {}
    if inis_path:
        meta = parse_inis(inis_path)
    print(f"[INFO] Device metadata: {meta}")
    
    device_config = {}
    if sups_path:
        device_config = parse_sups(sups_path)
    print(f"[INFO] Device config sections: {list(device_config.keys())}")
    
    # --- INHS (harmonics) ---
    _, inhs_df = parse_kew_file(inhs_path)
    if inhs_df is None:
        print("[ERROR] Could not parse INHS file.")
        return None
    print(f"[INFO] INHS parsed: {inhs_df.shape} rows")
    inhs_analysis = analyse_inhs(inhs_df)
    
    # --- INPS (interval power series) ---
    inps_analysis = {}
    if inps_path:
        _, inps_df = parse_inps(inps_path)
        if inps_df is not None:
            print(f"[INFO] INPS parsed: {inps_df.shape} rows")
            inps_analysis = analyse_inps(inps_df)
        else:
            print("[WARN] Could not parse INPS file.")
    
    # --- EVTS ---
    events = []
    if evts_path:
        events = parse_evts(evts_path)
    print(f"[INFO] Events found: {len(events)}")
    
    # --- VALS (RMS cycle-by-cycle tại sự kiện) ---
    vals_records = []
    if vals_path:
        vals_records = parse_vals(vals_path)
    
    # --- WAVS (dạng sóng raw tại sự kiện) ---
    wavs_records = []
    if wavs_path:
        wavs_records = parse_wavs(wavs_path)
    
    # --- Summary stats ---
    v_channels = [k for k in inhs_analysis if '[V]' in k and k.startswith('V')]
    a_channels = [k for k in inhs_analysis if '[A]' in k and k.startswith('A')]
    p_channels = [k for k in inhs_analysis if '[W]' in k]
    
    # INPS voltage/current/power dùng để bổ sung nếu INHS thiếu
    inps_v_cols = [k for k in inps_analysis if k.startswith('AVG_V') and '[V]' in k and not k.startswith('AVG_VL')]
    inps_a_cols = [k for k in inps_analysis if k.startswith('AVG_A') and '[A]' in k]
    inps_p_cols = [k for k in inps_analysis if k.startswith('AVG_P') and '[W]' in k]
    inps_q_cols = [k for k in inps_analysis if k.startswith('AVG_Q') and '[var]' in k]
    inps_s_cols = [k for k in inps_analysis if k.startswith('AVG_S') and '[VA]' in k]
    inps_pf_cols = [k for k in inps_analysis if k.startswith('AVG_PF') and '[_]' in k]
    inps_f_col = [k for k in inps_analysis if k.startswith('AVG_f') and '[Hz]' in k]
    inps_thd_v_cols = [k for k in inps_analysis if 'THDVR' in k]
    inps_thd_a_cols = [k for k in inps_analysis if 'THDAR' in k]
    
    def _inps_stat(col):
        d = inps_analysis.get(col, {})
        recorded_min = d.get('recorded_min')
        min_val = recorded_min if recorded_min is not None else d.get('min')
        recorded_max = d.get('recorded_max')
        max_val = recorded_max if recorded_max is not None else d.get('max')
        return {
            'avg': d.get('avg'),
            'min': min_val,
            'max': max_val,
        }
    
    summary = {
        'device': meta,
        'device_config': device_config,
        'folder': folder_name,
        'files': {
            'INHS': os.path.basename(inhs_path) if inhs_path else None,
            'INPS': os.path.basename(inps_path) if inps_path else None,
            'EVTS': os.path.basename(evts_path) if evts_path else None,
            'INIS': os.path.basename(inis_path) if inis_path else None,
            'SUPS': os.path.basename(sups_path) if sups_path else None,
            'VALS': os.path.basename(vals_path) if vals_path else None,
            'WAVS': os.path.basename(wavs_path) if wavs_path else None,
        },
        'time_start': inhs_analysis[v_channels[0]]['timestamps'][0] if v_channels else '',
        'time_end': inhs_analysis[v_channels[0]]['timestamps'][-1] if v_channels else '',
        'num_samples_inhs': len(inhs_analysis[v_channels[0]]['timestamps']) if v_channels else 0,
        'num_samples_inps': len(next(iter(inps_analysis.values()), {}).get('timestamps', [])),
        'channels': sorted(inhs_analysis.keys()),
        'events': events,
        'event_count': len(events),
        'voltage': {ch: {
            'avg': inhs_analysis[ch]['avg_fundamental'],
            'min': inhs_analysis[ch]['min_fundamental'],
            'max': inhs_analysis[ch]['max_fundamental'],
            'thd_avg': inhs_analysis[ch]['avg_thd'],
        } for ch in v_channels},
        'current': {ch: {
            'avg': inhs_analysis[ch]['avg_fundamental'],
            'min': inhs_analysis[ch]['min_fundamental'],
            'max': inhs_analysis[ch]['max_fundamental'],
            'thd_avg': inhs_analysis[ch]['avg_thd'],
        } for ch in a_channels},
        'power': {ch: {
            'avg': inhs_analysis[ch]['avg_fundamental'],
            'min': inhs_analysis[ch]['min_fundamental'],
            'max': inhs_analysis[ch]['max_fundamental'],
        } for ch in p_channels},
        # Dữ liệu từ INPS (đầy đủ hơn, bao gồm Q, S, PF, f, THD)
        'inps': {
            'voltage': {col: _inps_stat(col) for col in inps_v_cols},
            'current': {col: _inps_stat(col) for col in inps_a_cols},
            'active_power': {col: _inps_stat(col) for col in inps_p_cols},
            'reactive_power': {col: _inps_stat(col) for col in inps_q_cols},
            'apparent_power': {col: _inps_stat(col) for col in inps_s_cols},
            'power_factor': {col: _inps_stat(col) for col in inps_pf_cols},
            'frequency': {col: _inps_stat(col) for col in inps_f_col},
            'thd_voltage': {col: _inps_stat(col) for col in inps_thd_v_cols},
            'thd_current': {col: _inps_stat(col) for col in inps_thd_a_cols},
        } if inps_analysis else {},
    }
    
    return {
        'summary': summary,
        'series': inhs_analysis,
        'inps_series': inps_analysis,
        'vals': vals_records,    # RMS cycle-by-cycle tại event: [{channel, n_samples, values}]
        'wavs': wavs_records,    # Dạng sóng raw AC tại event: [{channel, n_samples, values}]
    }


def sanitize(obj):
    """Recursively replace nan/inf with None for JSON safety"""
    if isinstance(obj, float):
        if np.isnan(obj) or np.isinf(obj):
            return None
        return obj
    if isinstance(obj, list):
        return [sanitize(x) for x in obj]
    if isinstance(obj, dict):
        return {k: sanitize(v) for k, v in obj.items()}
    return obj


# ─────────────────────────────────────────────────────────
# Tạo nhận xét chất lượng điện tự động (tiếng Việt)
# ─────────────────────────────────────────────────────────

def generate_commentary(result, device_name='thiết bị'):
    """Tạo nhận xét chất lượng điện bằng tiếng Việt từ kết quả phân tích."""
    if not result:
        return "Không thể tạo nhận xét: dữ liệu phân tích không hợp lệ."

    summary = result.get('summary', {})
    series = result.get('series', {})
    inps = summary.get('inps', {})

    # --- Điện áp (nominal) ---
    nom_v_str = summary.get('device', {}).get('NOMINAL VOLTAGE', '400V')
    nom_v = float(re.sub(r'[^0-9.]', '', nom_v_str) or 400)
    # Nếu INHS đo điện áp pha trong khi nominal là điện áp dây
    nom_v_phase = nom_v / (3 ** 0.5) if nom_v > 300 else nom_v

    v_chs = ['V1[V]', 'V2[V]', 'V3[V]']
    a_chs = ['A1[A]', 'A2[A]', 'A3[A]']

    v_avgs = []
    for ch in v_chs:
        d = series.get(ch)
        if d:
            val = d.get('avg_fundamental') or 0
            if val > 0:
                v_avgs.append(val)

    if not v_avgs:
        return "Không thể tạo nhận xét: không có dữ liệu điện áp."

    v_mean_3ph = sum(v_avgs) / len(v_avgs)

    # Tính min/max điện áp pha từ chuỗi thời gian (lọc NaN và 0)
    v_series_min = {}
    v_series_max = {}
    for ch in v_chs:
        d = series.get(ch)
        if d:
            vals = [v for v in d.get('fundamental', []) if v is not None and v > 10]
            if vals:
                v_series_min[ch] = min(vals)
                v_series_max[ch] = max(vals)

    # Điện áp line-to-line: ưu tiên INPS VL (measured), fallback pha * √3
    inps_v = inps.get('voltage', {})
    vl_keys = sorted([k for k in inps_v if 'VL' in k])
    if vl_keys:
        # Có dữ liệu VL đo được từ INPS
        vl_avgs = [inps_v[k].get('avg') for k in vl_keys if inps_v[k] and inps_v[k].get('avg')]
        vl_mins_r = [inps_v[k].get('recorded_min') or inps_v[k].get('min') or inps_v[k].get('avg')
                     for k in vl_keys if inps_v[k]]
        vl_maxs_r = [inps_v[k].get('recorded_max') or inps_v[k].get('max') or inps_v[k].get('avg')
                     for k in vl_keys if inps_v[k]]
        vl_mins_r = [v for v in vl_mins_r if v is not None and v > 10]
        vl_maxs_r = [v for v in vl_maxs_r if v is not None and v > 10]
        v_line_min = min(vl_mins_r) if vl_mins_r else v_mean_3ph * (3**0.5)
        v_line_max = max(vl_maxs_r) if vl_maxs_r else v_mean_3ph * (3**0.5)
        v_line_nom = nom_v  # VL dùng nominal dây
    else:
        # Fallback: dùng chuỗi pha * √3 nếu mạng 3W (không có dây trung tính)
        wiring = summary.get('device', {}).get('WIRING', '')
        factor = (3**0.5) if '3W' in wiring else 1.0
        # Nếu v_mean_3ph đã là điện áp dây (>250V) thì không nhân thêm
        factor = (3**0.5) if v_mean_3ph < 260 else 1.0
        phase_mins = [v_series_min[ch] for ch in v_chs if ch in v_series_min]
        phase_maxs = [v_series_max[ch] for ch in v_chs if ch in v_series_max]
        v_line_min = min(phase_mins) * factor if phase_mins else v_mean_3ph * factor
        v_line_max = max(phase_maxs) * factor if phase_maxs else v_mean_3ph * factor
        v_line_nom = nom_v

    # δU: auto-detect nominal thực tế từ bảng tiêu chuẩn gần nhất với V_line đo được
    STANDARD_VOLTAGES = [110, 127, 220, 230, 240, 380, 400, 415, 440, 480, 600, 690, 1000]
    v_measured_avg = (v_line_min + v_line_max) / 2 if (v_line_min and v_line_max) else v_mean_3ph
    v_ref = min(STANDARD_VOLTAGES, key=lambda sv: abs(sv - v_measured_avg))
    # Nếu NOMINAL trong INIS gần với giá trị đo (±20%), ưu tiên
    if abs(nom_v - v_measured_avg) / max(nom_v, 1) < 0.20:
        v_ref = nom_v

    delta_u_min = (v_line_min - v_ref) / v_ref * 100
    delta_u_max = (v_line_max - v_ref) / v_ref * 100
    delta_u_ok = abs(delta_u_min) <= 5.0 and abs(delta_u_max) <= 5.0

    # Mất cân bằng điện áp ΔU (3 pha từ INHS)
    delta_U_imbalance = max(abs(v - v_mean_3ph) for v in v_avgs) / v_mean_3ph * 100 if v_mean_3ph > 0 else 0


    # --- Dòng điện ---
    a_avgs = []
    for ch in a_chs:
        d = series.get(ch)
        if d:
            val = d.get('avg_fundamental') or 0
            if val > 0:
                a_avgs.append(val)

    a_mean_3ph = sum(a_avgs) / len(a_avgs) if a_avgs else 1
    delta_I_imbalance = max(abs(a - a_mean_3ph) for a in a_avgs) / a_mean_3ph * 100 if a_avgs else 0

    # Biến động dòng điện
    a_all_vals = []
    for ch in a_chs:
        d = series.get(ch)
        if d:
            a_all_vals.extend([v for v in d.get('fundamental', []) if v is not None])

    if a_all_vals and a_mean_3ph > 0:
        spread = (max(a_all_vals) - min(a_all_vals)) / a_mean_3ph * 100
        if spread < 15:
            current_behavior = "ổn định trong thời gian đo kiểm"
        elif spread < 50:
            current_behavior = "biến đổi liên tục với biên độ nhỏ"
        else:
            current_behavior = "biến đổi liên tục trong thời gian đo kiểm"
    else:
        current_behavior = "biến đổi theo tải"

    # --- Hệ số công suất ---
    pf_val = None
    pf_inps = inps.get('power_factor', {})
    pf_total_key = next((k for k in pf_inps if k == 'AVG_PF[_]'), None)
    if pf_total_key and pf_inps[pf_total_key]:
        pf_val = pf_inps[pf_total_key].get('avg')

    if pf_val is None:
        va_chs = ['VA1[deg]', 'VA2[deg]', 'VA3[deg]']
        pf_list = [abs(float(np.cos(np.deg2rad(series[ch].get('avg_fundamental', 0) or 0))))
                   for ch in va_chs if series.get(ch)]
        pf_val = sum(pf_list) / len(pf_list) if pf_list else 0.0

    pf_abs = abs(pf_val or 0)
    if pf_abs >= 0.9:
        pf_level = "cao (trên 0,9)"
    elif pf_abs >= 0.8:
        pf_level = "trung bình (trên 0,8)"
    else:
        pf_level = "thấp (dưới 0,8)"

    # --- THD điện áp ---
    thd_v_list = [series[ch].get('avg_thd', 0) or 0 for ch in v_chs if series.get(ch)]
    thd_v_max = max(thd_v_list) if thd_v_list else 0
    # Bổ sung từ INPS THDVR (thường chính xác hơn INHS vì là avg over 1s)
    thd_v_inps = inps.get('thd_voltage', {})
    if thd_v_inps:
        extra = [v.get('max') or v.get('avg') or 0 for v in thd_v_inps.values() if v]
        if extra:
            thd_v_max = max(thd_v_max, max(v for v in extra if v is not None))

    # --- TDD dòng điện ---
    thd_a_list = [series[ch].get('avg_thd', 0) or 0 for ch in a_chs if series.get(ch)]
    tdd_max = max(thd_a_list) if thd_a_list else 0
    thd_a_inps = inps.get('thd_current', {})
    if thd_a_inps:
        extra = [v.get('max') or v.get('avg') or 0 for v in thd_a_inps.values() if v]
        if extra:
            tdd_max = max(tdd_max, max(v for v in extra if v is not None))

    thd_v_limit = 8.0
    tdd_limit = 20.0
    # Kiểm tra tính hợp lệ của các biến đo (tránh lỗi None)
    thd_v_ok = thd_v_max < thd_v_limit if thd_v_max is not None else True
    tdd_ok = tdd_max < tdd_limit if tdd_max is not None else True
    delta_u_ok_flag = delta_u_ok if delta_u_ok is not None else True
    delta_imb_ok = (delta_U_imbalance < 5.0 and delta_I_imbalance < 10.0) if (delta_U_imbalance is not None and delta_I_imbalance is not None) else True

    # Nếu hoàn toàn không có dữ liệu 3 pha, đánh giá là thiếu dữ liệu
    if not v_avgs:
        overall = "thiếu dữ liệu"
    else:
        overall = "tốt" if (delta_u_ok_flag and thd_v_ok and tdd_ok) else "cần cải thiện"

    # --- Format helpers ---
    def _f(v, d=1):
        if v is None:
            return '—'
        if isinstance(v, float) and np.isnan(v):
            return '—'
        return f"{v:.{d}f}".replace('.', ',')

    def _fp(v, d=2):
        if v is None:
            return '—'
        if isinstance(v, float) and np.isnan(v):
            return '—'
        sign = '+' if v > 0 else ''
        return f"{sign}{v:.{d}f}".replace('.', ',')

    commentary = (
        f"Nhận xét: Công suất tiêu thụ của máy biến áp tại {device_name} biến đổi và hoạt động ở mức {overall}. "
        f"Biểu đồ dòng điện tiêu thụ {current_behavior}. "
        f"Chất lượng điện đo tại {device_name} tương đối {'tốt' if overall == 'tốt' else 'kém'}, "
        f"hệ số công suất {pf_level}, "
        f"độ lệch pha {('thấp' if delta_imb_ok else 'cao')}, "
        f"tổng biến dạng sóng hài {('nhỏ' if thd_v_ok and tdd_ok else 'lớn')}. "
        f"Dưới đây là bảng tổng hợp thông số hoạt động:"
    )

    # --- Tạo Data Bảng ---
    table_rows = []

    # 1. Điện áp
    def get_v_stat(ch):
        d = series.get(ch)
        if not d: return None, None, None
        vals = [v for v in d.get('fundamental', []) if v is not None and v > 10]
        if not vals: return None, None, None
        return max(vals), min(vals), sum(vals)/len(vals)

    v1_max, v1_min, v1_avg = get_v_stat('V1[V]')
    v2_max, v2_min, v2_avg = get_v_stat('V2[V]')
    v3_max, v3_min, v3_avg = get_v_stat('V3[V]')

    if vl_keys and len(vl_keys) >= 3:
        u12_max, u12_min, u12_avg = inps_v[vl_keys[0]].get('recorded_max') or inps_v[vl_keys[0]].get('max'), inps_v[vl_keys[0]].get('recorded_min') or inps_v[vl_keys[0]].get('min'), inps_v[vl_keys[0]].get('avg')
        u23_max, u23_min, u23_avg = inps_v[vl_keys[1]].get('recorded_max') or inps_v[vl_keys[1]].get('max'), inps_v[vl_keys[1]].get('recorded_min') or inps_v[vl_keys[1]].get('min'), inps_v[vl_keys[1]].get('avg')
        u31_max, u31_min, u31_avg = inps_v[vl_keys[2]].get('recorded_max') or inps_v[vl_keys[2]].get('max'), inps_v[vl_keys[2]].get('recorded_min') or inps_v[vl_keys[2]].get('min'), inps_v[vl_keys[2]].get('avg')
    else:
        u12_max, u12_min, u12_avg = (v1_max * factor if v1_max else None), (v1_min * factor if v1_min else None), (v1_avg * factor if v1_avg else None)
        u23_max, u23_min, u23_avg = (v2_max * factor if v2_max else None), (v2_min * factor if v2_min else None), (v2_avg * factor if v2_avg else None)
        u31_max, u31_min, u31_avg = (v3_max * factor if v3_max else None), (v3_min * factor if v3_min else None), (v3_avg * factor if v3_avg else None)

    table_rows.append(["1", "Điện áp", "U12 (V)", _f(u12_max), _f(u12_min), _f(u12_avg), f"-5% ≤ δ ≤ 5% tại điện áp {int(v_ref)}V", "Đạt" if delta_u_ok else "Không đạt"])
    table_rows.append(["", "", "U23 (V)", _f(u23_max), _f(u23_min), _f(u23_avg), "", ""])
    table_rows.append(["", "", "U31 (V)", _f(u31_max), _f(u31_min), _f(u31_avg), "", ""])

    # 2. Dòng điện
    def get_a_stat(ch):
        d = series.get(ch)
        if not d: return None, None, None
        vals = [v for v in d.get('fundamental', []) if v is not None]
        if not vals: return None, None, None
        return max(vals), min(vals), sum(vals)/len(vals)

    i1_max, i1_min, i1_avg = get_a_stat('A1[A]')
    i2_max, i2_min, i2_avg = get_a_stat('A2[A]')
    i3_max, i3_min, i3_avg = get_a_stat('A3[A]')

    table_rows.append(["2", "Dòng điện", "I1 (A)", _f(i1_max), _f(i1_min), _f(i1_avg), "", ""])
    table_rows.append(["", "", "I2 (A)", _f(i2_max), _f(i2_min), _f(i2_avg), "", ""])
    table_rows.append(["", "", "I3 (A)", _f(i3_max), _f(i3_min), _f(i3_avg), "", ""])

    # 3. Độ lệch pha
    du_max = max(abs(delta_u_min), abs(delta_u_max))
    du_min = min(abs(delta_u_min), abs(delta_u_max))
    du_avg = (abs(delta_u_min) + abs(delta_u_max))/2
    table_rows.append(["3", "Độ lệch pha", "ΔU (%)", _f(du_max, 3), _f(du_min, 3), _f(du_avg, 3), "<5%", "Đạt" if delta_u_ok else "Không đạt"])
    table_rows.append(["", "", "ΔI (%)", _f(delta_I_imbalance, 3), _f(0, 3), _f(delta_I_imbalance/2, 3), "không quy định (< 10%)", "Đạt" if delta_I_imbalance < 10 else ""])

    # 4. Hệ số công suất PF
    pf_inps_item = inps.get('power_factor', {}).get(pf_total_key, {}) if pf_total_key else {}
    pf_max = pf_inps_item.get('max') or pf_val
    pf_min = pf_inps_item.get('min') or pf_val
    table_rows.append(["4", "Hệ số công suất", "PF", _f(pf_max, 3), _f(pf_min, 3), _f(pf_val, 3), "> 0,9", "Đã đạt" if (pf_val and pf_val >= 0.9) else "Cần lắp bù"])

    # 5. Công suất (P, Q, S total)
    def _get_p_stat(key_group, fallback_ch=None, scale=0.001):
        g = inps.get(key_group, {})
        # Ưu tiên các key tổng không chứa số pha 1, 2, 3 (ví dụ: AVG_P[W])
        total_keys = [k for k in g if 'AVG' in k and not any(ch in k for ch in ('1', '2', '3'))]
        k = total_keys[0] if total_keys else next((k for k in g if 'AVG' in k), None)
        item = g.get(k, {}) if k else {}
        
        # Nếu không có INPS, thử lấy từ INHS fallback
        f_max = item.get('max')
        f_min = item.get('min')
        f_avg = item.get('avg')
        
        if f_max is None and fallback_ch:
            f_max = series.get(fallback_ch, {}).get('max_fundamental')
        if f_min is None and fallback_ch:
            f_min = series.get(fallback_ch, {}).get('min_fundamental')
        if f_avg is None and fallback_ch:
            f_avg = series.get(fallback_ch, {}).get('avg_fundamental')
            
        return (f_max * scale if f_max is not None else None,
                f_min * scale if f_min is not None else None,
                f_avg * scale if f_avg is not None else None)

    p_max, p_min, p_avg = _get_p_stat('active_power', 'P[W]')
    q_max, q_min, q_avg = _get_p_stat('reactive_power')
    s_max, s_min, s_avg = _get_p_stat('apparent_power')

    table_rows.append(["5", "Công suất", "P (kW)", _f(p_max), _f(p_min), _f(p_avg), "", ""])
    table_rows.append(["", "", "Q (kVAR)", _f(q_max), _f(q_min), _f(q_avg), "", ""])
    table_rows.append(["", "", "S (kVA)", _f(s_max), _f(s_min), _f(s_avg), "", ""])

    # 6. THD điện áp
    def get_thd_stat(ch):
        d = series.get(ch)
        if not d or not d.get('thd'): return None, None, None
        vals = [v for v in d['thd'] if v is not None]
        if not vals: return None, None, None
        return max(vals), min(vals), sum(vals)/len(vals)

    thd_v1_max, thd_v1_min, thd_v1_avg = get_thd_stat('V1[V]')
    thd_v2_max, thd_v2_min, thd_v2_avg = get_thd_stat('V2[V]')
    thd_v3_max, thd_v3_min, thd_v3_avg = get_thd_stat('V3[V]')

    table_rows.append(["6", "Tổng biến dạng sóng\nhài điện áp", "THD1 (%)", _f(thd_v1_max,2), _f(thd_v1_min,2), _f(thd_v1_avg,2), "< 8%", "Đạt" if thd_v_ok else "Chưa đạt"])
    table_rows.append(["", "", "THD2 (%)", _f(thd_v2_max,2), _f(thd_v2_min,2), _f(thd_v2_avg,2), "", ""])
    table_rows.append(["", "", "THD3 (%)", _f(thd_v3_max,2), _f(thd_v3_min,2), _f(thd_v3_avg,2), "", ""])

    # 7. TDD dòng điện
    thd_a1_max, thd_a1_min, thd_a1_avg = get_thd_stat('A1[A]')
    thd_a2_max, thd_a2_min, thd_a2_avg = get_thd_stat('A2[A]')
    thd_a3_max, thd_a3_min, thd_a3_avg = get_thd_stat('A3[A]')

    table_rows.append(["7", "Tổng biến dạng sóng\nhài dòng điện", "TDD1 (%)", _f(thd_a1_max,2), _f(thd_a1_min,2), _f(thd_a1_avg,2), "<12%", "Đạt" if tdd_ok else "Chưa đạt"])
    table_rows.append(["", "", "TDD2 (%)", _f(thd_a2_max,2), _f(thd_a2_min,2), _f(thd_a2_avg,2), "", ""])
    table_rows.append(["", "", "TDD3 (%)", _f(thd_a3_max,2), _f(thd_a3_min,2), _f(thd_a3_avg,2), "", ""])

    return {
        "text": commentary,
        "table": table_rows
    }


# ─────────────────────────────────────────────────────────
# Xuất báo cáo ra file Excel hoàn chỉnh
# ─────────────────────────────────────────────────────────

def export_to_excel(commentary_obj, out_path="kew_analysis_report.xlsx"):
    try:
        if not commentary_obj or 'table' not in commentary_obj:
            return
        
        df = pd.DataFrame(commentary_obj['table'], columns=[
            "Stt", "Thông số", "Đại lượng đo", "Max", "Min", "Trung bình", "Quy chuẩn", "Nhận xét"
        ])
        
        with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Report', index=False, startrow=3)
            worksheet = writer.sheets['Report']
            worksheet.cell(row=1, column=1, value="BÁO CÁO CHẤT LƯỢNG ĐIỆN NĂNG")
            worksheet.cell(row=2, column=1, value=commentary_obj.get('text', ''))
            
            try:
                for column_cells in worksheet.columns:
                    length = max(len(str(cell.value) if cell.value is not None else "") for cell in column_cells)
                    col_letter = column_cells[0].column_letter
                    worksheet.column_dimensions[col_letter].width = min(length + 2, 60)
            except Exception:
                pass
        print(f"[OK] Saved complete report to {out_path}")
    except ImportError:
        print("[WARN] Pandas or openpyxl not installed, cannot export to Excel.")
    except Exception as e:
        print(f"[ERROR] Export to Excel failed: {e}")


if __name__ == '__main__':
    folder = sys.argv[1] if len(sys.argv) > 1 else 'S0224'
    result = build_analysis(folder)

    if result:
        out_file = 'kew_analysis.json'
        with open(out_file, 'w', encoding='utf-8') as f:
            json.dump(sanitize(result), f, ensure_ascii=False, indent=2, default=str)
        print(f"[OK] Saved analysis to {out_file}")
        
        commentary = generate_commentary(result, device_name=os.path.basename(folder))
        print(f"\n[NHẬN XÉT]\n{commentary['text']}")
        
        export_to_excel(commentary, out_path=f"kew_analysis_report_{os.path.basename(folder)}.xlsx")
    else:
        print("[ERROR] Analysis failed")
