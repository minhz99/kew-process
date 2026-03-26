import os
import zipfile
import io
import random
from PIL import Image, ImageDraw

TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), '../templates/digits')

# Caching templates
_digit_cache = {}
def get_digit_img(char, color='w'):
    key = f"{char}_{color}"
    if key in _digit_cache:
        return _digit_cache[key]
    
    char_map = {'.': 'dot', '-': 'minus', ':': 'colon'}
    filename_base = char_map.get(char, char)
    p = os.path.join(TEMPLATE_DIR, f"{filename_base}{color}.bmp")
    
    # Fallback to the other color if missing
    if not os.path.exists(p):
        other_color = 'g' if color == 'w' else 'w'
        p = os.path.join(TEMPLATE_DIR, f"{filename_base}{other_color}.bmp")
        
    if os.path.exists(p):
        img = Image.open(p).convert('RGB')
        _digit_cache[key] = img
        return img
        
    return None

def make_grid(ids, x_rights, y_bot, bg):
    res = []
    for i, id_val in enumerate(ids):
        res.append({'id': id_val, 'x': x_rights[i], 'y': y_bot, 'bg': bg})
    return res

# Screen configurations based on updated rules
sc0 = []
sc0.extend(make_grid(['V1','V2','V3'], [94,158,222], 54, 'w'))
sc0.extend(make_grid(['A1','A2','A3'], [94,158,222], 70, 'g'))
sc0.extend(make_grid(['P1','P2','P3'], [94,158,222], 86, 'w'))
sc0.extend(make_grid(['Q1','Q2','Q3'], [94,158,222], 102, 'g'))
sc0.extend(make_grid(['S1','S2','S3'], [94,158,222], 118, 'w'))
sc0.extend(make_grid(['PF1','PF2','PF3'], [94,158,222], 134, 'g'))
sc0.append({'id':'P', 'x':94, 'y':153, 'bg':'w'})
sc0.append({'id':'freq', 'alias':'f', 'x':222, 'y':153, 'bg':'w'})
sc0.append({'id':'Q', 'x':94, 'y':169, 'bg':'g'})
sc0.append({'id':'S', 'x':94, 'y':185, 'bg':'w'})
sc0.append({'id':'PF', 'x':94, 'y':201, 'bg':'g'})
sc0.append({'id':'An', 'x':222, 'y':201, 'bg':'g'})

sc1 = []
sc1.append({'id':'V1', 'x':69, 'y':38, 'bg':'w'})
sc1.append({'id':'V2', 'x':69, 'y':54, 'bg':'g'})
sc1.append({'id':'V3', 'x':69, 'y':70, 'bg':'w'})
sc1.append({'id':'A1', 'x':69, 'y':89, 'bg':'w'})
sc1.append({'id':'A2', 'x':69, 'y':105, 'bg':'g'})
sc1.append({'id':'A3', 'x':69, 'y':121, 'bg':'w'})
sc1.append({'id':'freq', 'alias':'f', 'x':89, 'y':156, 'bg':'w'})
sc1.append({'id':'V_unb', 'alias':'V%', 'x':89, 'y':191, 'bg':'g'})
sc1.append({'id':'A_unb', 'alias':'A%', 'x':89, 'y':207, 'bg':'w'})

sc2 = []
sc2.extend(make_grid(['V1','V2','V3'], [76,136,196], 47, 'w'))
sc2.extend(make_grid(['A1','A2','A3'], [76,136,196], 63, 'g'))

sc3 = []
sc3.extend(make_grid(['V1','V2','V3'], [76,136,196], 47, 'w'))
sc3.extend(make_grid(['A1','A2','A3'], [76,136,196], 63, 'g'))

sc4 = []
sc4.extend(make_grid(['V1','V2','V3'], [76,136,196], 47, 'w'))
sc4.extend(make_grid(['THDV1','THDV2','THDV3'], [76,136,196], 63, 'g'))

sc5 = []
sc5.extend(make_grid(['A1','A2','A3'], [76,136,196], 47, 'w'))
sc5.extend(make_grid(['THDA1','THDA2','THDA3'], [76,136,196], 63, 'g'))

SCREENS = [
    {'id':'SD140', 'overlays':sc0},
    {'id':'SD141', 'overlays':sc1},
    {'id':'SD142', 'overlays':sc2},
    {'id':'SD143', 'overlays':sc3},
    {'id':'SD144', 'overlays':sc4},
    {'id':'SD145', 'overlays':sc5}
]

def apply_text_to_image(img, overlay_config, text_str):
    x_right = overlay_config['x']
    y_bot = overlay_config['y']
    color = overlay_config.get('bg', 'w')
    
    # 50x15 clear space right aligned
    w_clear = 50
    h_clear = 15
    x_left = x_right - w_clear + 1
    y_top = y_bot - h_clear + 1
    
    if x_left < 0: x_left = 0
    if y_top < 0: y_top = 0
    if x_right >= img.width: x_right = img.width - 1
    if y_bot >= img.height: y_bot = img.height - 1
    
    # 1. Clear background
    bg_color = img.getpixel((x_left, y_bot))
    draw = ImageDraw.Draw(img)
    draw.rectangle([x_left, y_top, x_right, y_bot], fill=bg_color)
    
    # 2. Draw new text right-aligned
    chars = list(text_str)
    curr_x = x_right + 1
    
    for char in reversed(chars):
        if char == '/': char = '.'
        digit_img = get_digit_img(char, color)
        if digit_img:
            w_digit = digit_img.width
            spacing = 1 if w_digit == 9 else 2
            
            curr_x -= w_digit
            paste_y = y_bot - digit_img.height + 1
            
            # Using paste directly, ignoring transparency since background is matched
            img.paste(digit_img, (curr_x, paste_y))
            
            curr_x -= spacing
        else:
            curr_x -= 6

def process_image(img, screen_idx, params):
    sc = SCREENS[screen_idx % 6]
    
    for overlay in sc['overlays']:
        # Try finding parameter by ID or its alias
        val = params.get(overlay['id'])
        if val is None and 'alias' in overlay:
            val = params.get(overlay['alias'])
            
        if val is not None:
            apply_text_to_image(img, overlay, str(val))
            
    return img

def randomize_val(val_str, fluctuation=0.005):
    try:
        v = float(val_str)
        delta = v * fluctuation * random.uniform(-1, 1)
        if '.' in val_str:
            decimals = len(val_str.split('.')[1])
            return f"{v + delta:.{decimals}f}"
        else:
            return f"{int(v + delta)}"
    except:
        return val_str

def process_zip(zip_bytes, params):
    in_zip = zipfile.ZipFile(zip_bytes, 'r')
    out_bytes = io.BytesIO()
    out_zip = zipfile.ZipFile(out_bytes, 'w')
    
    files = sorted([f for f in in_zip.namelist() if f.upper().endswith('.BMP')])
    
    fluctuate = str(params.get('fluctuate', 'false')).lower() == 'true'
    
    processed_count = 0
    for idx, filename in enumerate(files):
        try:
            img_data = in_zip.read(filename)
            img = Image.open(io.BytesIO(img_data)).convert('RGB')
            
            img_params = {}
            for k, v in params.items():
                if fluctuate and k not in ['fluctuate', 'date', 'time']:
                    img_params[k] = randomize_val(v)
                else:
                    img_params[k] = v
                    
            img = process_image(img, idx, img_params)
            
            img_out_bytes = io.BytesIO()
            img.save(img_out_bytes, format='BMP')
            out_zip.writestr(filename, img_out_bytes.getvalue())
            processed_count += 1
        except Exception as e:
            print(f"[Error] Failed to process {filename}: {e}")
            continue
            
    out_zip.close()
    if processed_count == 0:
        raise Exception("Không tìm thấy file ảnh .BMP hợp lệ nào trong tệp ZIP.")
    out_bytes.seek(0)
    return out_bytes
