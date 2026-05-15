import pandas as pd
import json
import os

CONFIG = {}
config_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'config.json')
if os.path.exists(config_path):
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            CONFIG = json.load(f)
    except Exception:
        pass

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

def find_file(folder, prefix):
    for f in os.listdir(folder):
        if f.upper().startswith(prefix.upper()) and f.upper().endswith('.KEW'):
            return os.path.join(folder, f)
    return None



