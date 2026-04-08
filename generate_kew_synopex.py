"""
generate_kew_synopex.py
=======================
Tạo báo cáo KEW tự động — Dự án SYNOPEX BẮC NINH 335-357.
- Clone file mẫu giữ nguyên 100% styles / font / page setup.
- Tên máy tự động lấy từ tên thư mục (dạng "S0335 - Tên máy").
- Dùng Tesseract OCR đọc thông số từ ảnh để tự điền Nhận xét.

Yêu cầu:
    1. Cài Tesseract: https://github.com/UB-Mannheim/tesseract/wiki
    2. pip install pytesseract pillow lxml

Chạy:   python generate_kew_synopex.py
"""

import os, re, shutil, zipfile, copy, traceback
from lxml import etree
from PIL import Image

try:
    import pytesseract
except ImportError:
    pytesseract = None

from modules.image.kew6315_ocr import coerce_number, read_kew6315_screen_fields

# ==============================================================================
# CẤU HÌNH — CHỈ CẦN CHỈNH 4 DÒNG NÀY
# ==============================================================================
TEMPLATE_FILE = r"C:\Polytee\file_mau_kew.docx"
BASE_DIR      = r"C:\Polytee\KEW SYNOPEX BẮC NINH 335-357\KEW SYNOPEX BẮC NINH 335-357"
OUTPUT_FILE   = r"C:\Polytee\KEW_Synopex_BacNinh_Report.docx"
TESSERACT_CMD = r"C:\Users\Acer\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
# ==============================================================================

if pytesseract is not None:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

VALID_EXT = ('.png', '.jpg', '.jpeg', '.bmp')

# Kích thước ảnh EMU (từ file mẫu KEW chuẩn)
TREND_CX  = "6048000";  TREND_CY  = "2012400"   # 6.61" x 2.20"
DETAIL_CX = "1990800";  DETAIL_CY = "1486800"   # 2.18" x 1.63"

# Namespaces Word
W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

def qw(tag): return f"{{{W}}}{tag}"


# ==============================================================================
# PARSE TÊN MÁY TỪ TÊN THƯ MỤC
# ==============================================================================

def parse_folder(folder_name):
    """
    Tên thư mục dạng: "S0335 - Tủ bơm nước sinh hoạt"
    → ma_so = "335"
    → ten_may = "Tủ bơm nước sinh hoạt"

    Nếu không có dấu " - " → dùng tên folder làm tên máy.
    """
    m = re.match(r'^[Ss](\d+)\s*-\s*(.+)$', folder_name.strip())
    if m:
        ma_so  = m.group(1).lstrip('0') or m.group(1)  # bỏ số 0 đầu: 0335 → 335
        ten_may = m.group(2).strip()
        return ma_so, ten_may
    # Fallback: lấy số từ tên thư mục
    m2 = re.match(r'^[Ss](\d+)', folder_name)
    ma_so = m2.group(1).lstrip('0') if m2 else folder_name
    return ma_so, folder_name


def tc_title(name):
    """TC → Tổng cấp (cho tiêu đề và caption)."""
    if name.startswith("TC "):
        return "Tổng cấp " + name[3:]
    return name


def expand_tc(name, first=True):
    """TC đầu tiên → 'Tổng cấp', các lần sau → 'tổng cấp'."""
    if not name.startswith("TC "):
        return name
    return ("Tổng cấp " if first else "tổng cấp ") + name[3:]


# ==============================================================================
# OCR — Ưu tiên nhận diện theo template pixel, fallback về Tesseract khi cần
# ==============================================================================

def crop(img, region):
    """Crop ảnh theo tỉ lệ (x1,y1,x2,y2) từ 0.0 đến 1.0."""
    w, h = img.size
    return img.crop((int(region[0]*w), int(region[1]*h),
                     int(region[2]*w), int(region[3]*h)))

def ocr_text(img_pil):
    """Tesseract --psm 6, resize x2, không whitelist."""
    if pytesseract is None:
        raise RuntimeError("pytesseract chưa được cài đặt")
    w, h = img_pil.size
    img_big = img_pil.resize((w*2, h*2), Image.LANCZOS)
    return pytesseract.image_to_string(img_big, lang="eng", config="--psm 6")


def _read_template_fields(img_path, screen_idx, field_ids):
    try:
        fields = read_kew6315_screen_fields(img_path, screen_idx=screen_idx, field_ids=field_ids)
        printable = {k: v for k, v in fields.items() if v not in (None, "")}
        if printable:
            print(f"        Screen {screen_idx} template: {printable}")
        return fields
    except Exception as e:
        print(f"        [Template OCR ERR screen {screen_idx}] {e}")
        return {}


def _pick_values(fields, field_ids, min_value=None, max_value=None):
    values = []
    for field_id in field_ids:
        value = coerce_number(fields.get(field_id))
        if value is None:
            continue
        if min_value is not None and value < min_value:
            continue
        if max_value is not None and value > max_value:
            continue
        values.append(value)
    return values


def _read_h2_tesseract(img_path):
    img = Image.open(img_path).convert("RGB")

    c_vl = crop(img, (0.00, 0.167, 0.80, 0.250))
    text_vl = ocr_text(c_vl)
    print(f"        H2 VL raw (fallback): {repr(text_vl.strip()[:100])}")
    vl_nums = re.findall(r'\d+\.\d+|\d+', text_vl)
    vl_vals = [float(v) for v in vl_nums if 300 <= float(v) <= 500][:3]

    c_pf = crop(img, (0.00, 0.500, 0.70, 0.583))
    text_pf = ocr_text(c_pf)
    print(f"        H2 PF raw (fallback): {repr(text_pf.strip()[:100])}")
    pf_nums = re.findall(r'\d+\.\d+|\d+', text_pf)
    pf_vals = [float(v) for v in pf_nums if 0.50 <= float(v) <= 1.00][:3]
    return vl_vals, pf_vals


def read_h2(img_path):
    """
    Hình 2 (ps-sd index 0) — màn hình SD140.
    Đọc trực tiếp V1/V2/V3 và PF1/PF2/PF3 theo đúng tọa độ field.
    """
    try:
        fields = _read_template_fields(img_path, 0, ["V1", "V2", "V3", "PF1", "PF2", "PF3"])
        vl_vals = _pick_values(fields, ["V1", "V2", "V3"], 300.0, 500.0)
        pf_vals = _pick_values(fields, ["PF1", "PF2", "PF3"], 0.50, 1.00)

        if not vl_vals or not pf_vals:
            fallback_vl, fallback_pf = _read_h2_tesseract(img_path)
            if not vl_vals:
                vl_vals = fallback_vl
            if not pf_vals:
                pf_vals = fallback_pf

        vl_min = f"{min(vl_vals):.1f}" if vl_vals else "..."
        vl_max = f"{max(vl_vals):.1f}" if vl_vals else "..."
        pf = max(pf_vals) if pf_vals else None
        return vl_min, vl_max, pf

    except Exception as e:
        print(f"        [H2 ERR] {e}")
        return "...", "...", None


def _read_h3_tesseract(img_path):
    img = Image.open(img_path).convert("RGB")

    c_v = crop(img, (0.10, 0.74, 0.42, 0.83))
    text_v = ocr_text(c_v)
    print(f"        H3 V raw (fallback): {repr(text_v.strip()[:80])}")
    nums_v = re.findall(r'\d+\.\d+|\d+', text_v)
    vals_v = [float(v) for v in nums_v if 0.0 <= float(v) <= 20.0]

    c_a = crop(img, (0.10, 0.80, 0.42, 0.89))
    text_a = ocr_text(c_a)
    print(f"        H3 A raw (fallback): {repr(text_a.strip()[:80])}")
    nums_a = re.findall(r'\d+\.\d+|\d+', text_a)
    vals_a = [float(v) for v in nums_a if 0.0 <= float(v) <= 20.0]
    return vals_v, vals_a


def read_h3(img_path):
    """
    Hình 3 (ps-sd index 1) — màn hình SD141.
    Đọc trực tiếp V_unb và A_unb theo field cố định.
    """
    try:
        fields = _read_template_fields(img_path, 1, ["V_unb", "A_unb"])
        vals_v = _pick_values(fields, ["V_unb"], 0.0, 20.0)
        vals_a = _pick_values(fields, ["A_unb"], 0.0, 20.0)

        if not vals_v or not vals_a:
            fallback_v, fallback_a = _read_h3_tesseract(img_path)
            if not vals_v:
                vals_v = fallback_v
            if not vals_a:
                vals_a = fallback_a

        delta_u = f"{vals_v[0]:.1f}" if vals_v else "..."
        delta_i = f"{vals_a[0]:.1f}" if vals_a else "..."
        return delta_u, delta_i

    except Exception as e:
        print(f"        [H3 ERR] {e}")
        return "...", "..."


def _read_thd_tesseract(img_path):
    img = Image.open(img_path).convert("RGB")
    c = crop(img, (0.00, 0.20, 0.70, 0.30))
    text = ocr_text(c)
    print(f"        THD raw (fallback): {repr(text.strip()[:80])}")
    nums = re.findall(r'\d+\.\d+|\d+', text)
    return [float(v) for v in nums if 0.0 < float(v) < 100.0][:3]


def read_thd(img_path, screen_idx, field_ids):
    """
    Hình 6/7 (ps-sd index 4/5) — màn hình SD144/SD145.
    Đọc trực tiếp 3 field sóng hài và lấy max.
    """
    try:
        fields = _read_template_fields(img_path, screen_idx, field_ids)
        vals = _pick_values(fields, field_ids, 0.0, 100.0)

        if not vals:
            vals = _read_thd_tesseract(img_path)

        return f"{max(vals):.2f}" if vals else "..."

    except Exception as e:
        print(f"        [THD ERR] {e}")
        return "..."


def read_all_params(detail_paths):
    """Đọc thông số từ 6 ảnh detail. index: 0=H2, 1=H3, 4=H6, 5=H7."""
    result = {
        "vl_min": "...", "vl_max": "...", "pf": None,
        "delta_u": "...", "delta_i": "...",
        "thd_max": "...", "tdd_max": "..."
    }
    try:
        vl_min, vl_max, pf = read_h2(detail_paths[0])
        result["vl_min"] = vl_min
        result["vl_max"] = vl_max
        result["pf"]     = pf
    except Exception as e:
        print(f"        [H2 ERR] {e}")

    try:
        du, di = read_h3(detail_paths[1])
        result["delta_u"] = du
        result["delta_i"] = di
    except Exception as e:
        print(f"        [H3 ERR] {e}")

    try:
        result["thd_max"] = read_thd(detail_paths[4], 4, ["THDV1", "THDV2", "THDV3"])
    except Exception as e:
        print(f"        [H6 ERR] {e}")

    try:
        result["tdd_max"] = read_thd(detail_paths[5], 5, ["THDA1", "THDA2", "THDA3"])
    except Exception as e:
        print(f"        [H7 ERR] {e}")

    return result


# ==============================================================================
# TÍNH TOÁN VÀ XÂY DỰNG NHẬN XÉT
# ==============================================================================

def pf_level(pf):
    if pf is None: return "..."
    return "cao (trên 0,8)" if pf >= 0.8 else "thấp (dưới 0,8)"

def calc_du(vl_min_str, vl_max_str):
    try:
        du_min = (float(vl_min_str) - 400.0) / 4.0
        du_max = (float(vl_max_str) - 400.0) / 4.0
        return f"{du_min:+.2f}", f"{du_max:+.2f}"
    except:
        return "...", "..."

def du_status(du_min_str, du_max_str):
    try:
        ok = -5.0 <= float(du_min_str) <= 5.0 and -5.0 <= float(du_max_str) <= 5.0
        return "đạt tiêu chuẩn" if ok else "không đạt tiêu chuẩn"
    except:
        return "đạt tiêu chuẩn"

def phase_status(delta_u, delta_i):
    try:
        ok = float(delta_u) < 5.0 and float(delta_i) < 10.0
        return "đều ở mức thấp" if ok else "ở mức cao"
    except:
        return "đều ở mức thấp"

def build_thd_sentence(thd_str, tdd_str):
    """4 trường hợp kết luận THD/TDD theo chuẩn file mẫu."""
    try:
        thd = float(thd_str)
        tdd = float(tdd_str)
        thd_ok = thd < 8.0
        tdd_ok = tdd < 20.0
        thd_fmt = f"{thd:.2f}".replace('.', ',')
        tdd_fmt = f"{tdd:.2f}".replace('.', ',')

        if thd_ok and tdd_ok:
            return (
                f"Tổng biến dạng sóng hài điện áp và dòng điện đều ở mức cho phép "
                f"(THDmax = {thd_fmt}% < 8,0% & TDDmax = {tdd_fmt}% < 20,0%)."
            )
        elif thd_ok and not tdd_ok:
            return (
                f"Tổng biến dạng sóng hài điện áp ở mức cho phép "
                f"(THDmax = {thd_fmt}% < 8,0%); tuy nhiên, tổng biến dạng sóng hài "
                f"dòng điện cao hơn mức cho phép (TDDmax = {tdd_fmt}% > 20,0%)."
            )
        elif not thd_ok and tdd_ok:
            return (
                f"Tổng biến dạng sóng hài dòng điện ở mức cho phép "
                f"(TDDmax = {tdd_fmt}% < 20,0%); tuy nhiên, tổng biến dạng sóng hài "
                f"điện áp cao hơn mức cho phép (THDmax = {thd_fmt}% > 8,0%)."
            )
        else:
            return (
                f"Tổng biến dạng sóng hài điện áp và dòng điện đều vượt mức cho phép "
                f"(THDmax = {thd_fmt}% > 8,0% & TDDmax = {tdd_fmt}% > 20,0%)."
            )
    except:
        return (
            f"Tổng biến dạng sóng hài điện áp và dòng điện đều ở mức cho phép "
            f"(THDmax = {thd_str}% < 8,0% & TDDmax = {tdd_str}% < 20,0%)."
        )

def build_nhanxet(ma_so, ten_may, p):
    du_min, du_max = calc_du(p["vl_min"], p["vl_max"])

    def fmt(s):
        try: return f"{float(s):.1f}".replace('.', ',')
        except: return s

    vl_min_fmt = fmt(p["vl_min"])
    vl_max_fmt = fmt(p["vl_max"])
    du_min_fmt = du_min.replace('.', ',') if du_min != "..." else "..."
    du_max_fmt = du_max.replace('.', ',') if du_max != "..." else "..."
    du_ok_txt  = du_status(du_min, du_max)
    dv_fmt     = fmt(p["delta_u"])
    di_fmt     = fmt(p["delta_i"])

    ten_first = expand_tc(ten_may, first=True)

    return (
        f"Biểu đồ dòng điện tiêu thụ tại {ten_first} biến đổi... "
        f"trong thời gian đo kiểm. "
        f"Hệ số công suất cosφ ở mức {pf_level(p['pf'])}. "
        f"Điện áp dao động từ {vl_min_fmt} ÷ {vl_max_fmt} V, "
        f"độ lệch chuẩn của điện áp δU (= {du_min_fmt}% ÷ {du_max_fmt}%) "
        f"{du_ok_txt} (-5,0% ≤ δ ≤ 5,0%). "
        f"Độ lệch pha điện áp và dòng điện {phase_status(p['delta_u'], p['delta_i'])} "
        f"(ΔU = {dv_fmt}% < 5,0%, ΔI = {di_fmt}% < 10,0%). "
        + build_thd_sentence(p["thd_max"], p["tdd_max"])
    )


# ==============================================================================
# BUILDER — Clone template & inject content
# ==============================================================================

class KewReportBuilder:

    def __init__(self):
        self.work_dir    = OUTPUT_FILE + "_workdir"
        self.media_dir   = os.path.join(self.work_dir, "word", "media")
        self._rid_num    = 200
        self._img_num    = 1
        self._docpr_num  = 900000000
        self.fig_counter = 1

    # ── Setup ──────────────────────────────────────────────────────────────────

    def unpack(self):
        if os.path.exists(self.work_dir):
            shutil.rmtree(self.work_dir)
        with zipfile.ZipFile(TEMPLATE_FILE, 'r') as z:
            z.extractall(self.work_dir)
        print("✓ Giải nén file mẫu")

    def load(self):
        self.doc_path  = os.path.join(self.work_dir, "word", "document.xml")
        self.rels_path = os.path.join(self.work_dir, "word", "_rels", "document.xml.rels")
        self.ct_path   = os.path.join(self.work_dir, "[Content_Types].xml")

        self.doc_tree  = etree.parse(self.doc_path)
        self.rels_tree = etree.parse(self.rels_path)
        self.ct_tree   = etree.parse(self.ct_path)

        self.body      = self.doc_tree.getroot().find(f".//{{{W}}}body")
        self.rels_root = self.rels_tree.getroot()
        self.ct_root   = self.ct_tree.getroot()

        for rel in self.rels_root:
            m = re.search(r"rId(\d+)", rel.get("Id", ""))
            if m: self._rid_num = max(self._rid_num, int(m.group(1)) + 1)

        if os.path.exists(self.media_dir):
            for f in os.listdir(self.media_dir):
                m = re.search(r"image(\d+)", f)
                if m: self._img_num = max(self._img_num, int(m.group(1)) + 1)

        print(f"✓ Load XML (rId từ rId{self._rid_num}, image từ image{self._img_num})")

    def clear_body(self):
        sect = self.body.find(qw("sectPr"))
        sect_copy = copy.deepcopy(sect) if sect is not None else None
        for child in list(self.body): self.body.remove(child)
        if sect_copy is not None: self.body.append(sect_copy)
        print("✓ Xóa body cũ")

    # ── Image ──────────────────────────────────────────────────────────────────

    def _add_image(self, src_path):
        ext      = os.path.splitext(src_path)[1].lower().lstrip(".")
        img_name = f"image{self._img_num}.{ext}"
        shutil.copy2(src_path, os.path.join(self.media_dir, img_name))
        rid = f"rId{self._rid_num}"
        self._img_num += 1; self._rid_num += 1

        rel = etree.SubElement(self.rels_root, "Relationship")
        rel.set("Id", rid); rel.set("Type", REL_IMAGE)
        rel.set("Target", f"media/{img_name}")

        mime = {"png":"image/png","jpg":"image/jpeg","jpeg":"image/jpeg","bmp":"image/bmp"}.get(ext, f"image/{ext}")
        if ext not in {el.get("Extension","").lower() for el in self.ct_root}:
            e = etree.SubElement(self.ct_root, "Default")
            e.set("Extension", ext); e.set("ContentType", mime)
        return rid

    def _pid(self):
        self._docpr_num += 1
        return str(self._docpr_num)

    # ── XML builders ───────────────────────────────────────────────────────────

    def _drawing(self, rid, cx, cy):
        p = self._pid()
        return f"""<w:drawing xmlns:w="{W}" xmlns:wp="{WP}" xmlns:a="{A}" xmlns:pic="{PIC}" xmlns:r="{R}">
  <wp:inline distT="0" distB="0" distL="0" distR="0">
    <wp:extent cx="{cx}" cy="{cy}"/>
    <wp:effectExtent l="0" t="0" r="0" b="0"/>
    <wp:docPr id="{p}" name="img{p}"/>
    <wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>
    <a:graphic>
      <a:graphicData uri="{PIC}">
        <pic:pic>
          <pic:nvPicPr>
            <pic:cNvPr id="0" name="img{p}"/>
            <pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr>
          </pic:nvPicPr>
          <pic:blipFill>
            <a:blip r:embed="{rid}"/>
            <a:stretch><a:fillRect/></a:stretch>
          </pic:blipFill>
          <pic:spPr bwMode="auto">
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:noFill/><a:ln><a:noFill/></a:ln>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>"""

    def _cpara(self, drw, align="center"):
        return f"""<w:p xmlns:w="{W}">
  <w:pPr><w:pStyle w:val="KHnhtrongbng"/><w:jc w:val="{align}"/>
    <w:rPr><w:sz w:val="24"/></w:rPr></w:pPr>
  <w:r><w:rPr><w:noProof/><w:sz w:val="24"/></w:rPr>{drw}</w:r>
</w:p>"""

    def _tc(self, para, span=None):
        sp = f'<w:gridSpan w:val="{span}"/>' if span else ""
        return f"""<w:tc xmlns:w="{W}">
  <w:tcPr><w:tcW w:w="1666" w:type="pct"/>{sp}
    <w:tcBorders>
      <w:top    w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:left   w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:right  w:val="none" w:sz="0" w:space="0" w:color="auto"/>
    </w:tcBorders>
    <w:vAlign w:val="center"/>
  </w:tcPr>
  {para}
</w:tc>"""

    def _row(self, cells):
        return f'<w:tr xmlns:w="{W}"><w:trPr><w:jc w:val="center"/></w:trPr>{cells}</w:tr>'

    # ── Block builders ─────────────────────────────────────────────────────────

    def make_title(self, ma_so, ten_may):
        return etree.fromstring(f"""<w:p xmlns:w="{W}">
  <w:pPr><w:pStyle w:val="onvn"/></w:pPr>
  <w:r><w:t xml:space="preserve">{tc_title(ten_may)}:</w:t></w:r>
</w:p>""")

    def make_table(self, folder_path, trend_img, detail_imgs):
        tr = self._add_image(os.path.join(folder_path, trend_img))
        dr = [self._add_image(os.path.join(folder_path, d)) for d in detail_imgs[:6]]
        tp = self._cpara(self._drawing(tr, TREND_CX, TREND_CY), "left")
        dp = [self._cpara(self._drawing(r, DETAIL_CX, DETAIL_CY)) for r in dr]
        return etree.fromstring(f"""<w:tbl xmlns:w="{W}">
  <w:tblPr>
    <w:tblW w:w="5000" w:type="pct"/><w:jc w:val="center"/>
    <w:tblLayout w:type="fixed"/>
    <w:tblBorders>
      <w:top     w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:left    w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:bottom  w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:right   w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>
    </w:tblBorders>
    <w:tblCellMar>
      <w:left w:w="115" w:type="dxa"/><w:right w:w="115" w:type="dxa"/>
    </w:tblCellMar>
  </w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="3216"/><w:gridCol w:w="3216"/><w:gridCol w:w="3220"/>
  </w:tblGrid>
  {self._row(self._tc(tp, span=3))}
  {self._row(self._tc(dp[0]) + self._tc(dp[1]) + self._tc(dp[2]))}
  {self._row(self._tc(dp[3]) + self._tc(dp[4]) + self._tc(dp[5]))}
</w:tbl>""")

    def make_caption(self, fig_num, ma_so, ten_may):
        return etree.fromstring(f"""<w:p xmlns:w="{W}">
  <w:pPr><w:pStyle w:val="DMhnhK"/></w:pPr>
  <w:r><w:t xml:space="preserve">Hình 4.{fig_num} - Kết quả đo chất lượng điện {tc_title(ten_may)}</w:t></w:r>
</w:p>""")

    def make_nhanxet(self, text):
        safe = text.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
        return etree.fromstring(f"""<w:p xmlns:w="{W}">
  <w:pPr>
    <w:pStyle w:val="onvnnhnxt"/>
    <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
    <w:jc w:val="both"/>
  </w:pPr>
  <w:r><w:rPr><w:b/><w:u w:val="single"/></w:rPr><w:t>Nhận xét:</w:t></w:r>
  <w:r><w:t xml:space="preserve"> </w:t></w:r>
  <w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">{safe}</w:t></w:r>
</w:p>""")

    def make_pagebreak(self):
        return etree.fromstring(f"""<w:p xmlns:w="{W}">
  <w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
    <w:rPr><w:i/></w:rPr></w:pPr>
  <w:r><w:rPr><w:i/></w:rPr><w:br w:type="page"/></w:r>
</w:p>""")

    # ── Inject ─────────────────────────────────────────────────────────────────

    def inject(self, ma_so, ten_may, folder_path, trend_img, detail_imgs, nhanxet):
        sect = self.body.find(qw("sectPr"))
        els  = [
            self.make_title(ma_so, ten_may),
            self.make_table(folder_path, trend_img, detail_imgs),
            self.make_caption(self.fig_counter, ma_so, ten_may),
            self.make_nhanxet(nhanxet),
            self.make_pagebreak(),
        ]
        if sect is not None:
            idx = list(self.body).index(sect)
            for i, el in enumerate(els): self.body.insert(idx + i, el)
        else:
            for el in els: self.body.append(el)
        self.fig_counter += 1

    # ── Save ───────────────────────────────────────────────────────────────────

    def save(self):
        self.doc_tree.write( self.doc_path,  xml_declaration=True, encoding="UTF-8", standalone=True)
        self.rels_tree.write(self.rels_path, xml_declaration=True, encoding="UTF-8", standalone=True)
        self.ct_tree.write(  self.ct_path,   xml_declaration=True, encoding="UTF-8", standalone=True)

        if os.path.exists(OUTPUT_FILE):
            try: os.remove(OUTPUT_FILE)
            except PermissionError:
                print("✗ File đang mở trong Word — hãy đóng lại rồi chạy lại!")
                shutil.rmtree(self.work_dir, ignore_errors=True); return

        with zipfile.ZipFile(OUTPUT_FILE, 'w', zipfile.ZIP_DEFLATED) as z:
            for root, dirs, files in os.walk(self.work_dir):
                for f in files:
                    full = os.path.join(root, f)
                    z.write(full, os.path.relpath(full, self.work_dir))

        shutil.rmtree(self.work_dir)
        print(f"\n✓ Hoàn thành! Lưu tại: {OUTPUT_FILE}")

    # ── Main ───────────────────────────────────────────────────────────────────

    def build(self):
        if pytesseract is None or not os.path.exists(TESSERACT_CMD):
            print("! Tesseract fallback không sẵn sàng, sẽ chỉ dùng OCR theo template pixel.")
            if pytesseract is None:
                print("  → Thiếu module `pytesseract`.")
            else:
                print(f"  → Không tìm thấy Tesseract: {TESSERACT_CMD}")

        if not os.path.exists(TEMPLATE_FILE):
            print(f"✗ Không tìm thấy file mẫu: {TEMPLATE_FILE}"); return

        if not os.path.exists(BASE_DIR):
            print(f"✗ Không tìm thấy thư mục ảnh: {BASE_DIR}"); return

        # Lấy danh sách thư mục S* (bao gồm cả tên có khoảng trắng như "S0335 - Tủ bơm...")
        subdirs = sorted([
            d for d in os.listdir(BASE_DIR)
            if os.path.isdir(os.path.join(BASE_DIR, d)) and re.match(r'^[Ss]\d+', d)
        ])
        if not subdirs:
            print(f"✗ Không tìm thấy thư mục S* trong {BASE_DIR}"); return

        print(f"✓ Tìm thấy {len(subdirs)} thư mục máy")
        self.unpack(); self.load(); self.clear_body()

        ok = skip = 0
        for folder_name in subdirs:
            folder_path = os.path.join(BASE_DIR, folder_name)
            files       = os.listdir(folder_path)

            trend_imgs  = sorted([f for f in files
                                   if f.lower().startswith("a") and f.lower().endswith(VALID_EXT)])
            detail_imgs = sorted([f for f in files
                                   if f.lower().startswith("ps-sd") and f.lower().endswith(VALID_EXT)])

            if not trend_imgs or len(detail_imgs) < 6:
                print(f"  [bỏ qua] {folder_name} — cần 1 ảnh 'a' và ít nhất 6 ảnh 'ps-sd'.")
                skip += 1; continue

            # Lấy tên máy tự động từ tên thư mục
            ma_so, ten_may = parse_folder(folder_name)
            print(f"\n  ── {folder_name}")
            print(f"     → Máy {ma_so}: {ten_may}")

            # OCR đọc thông số
            detail_paths = [os.path.join(folder_path, d) for d in detail_imgs[:6]]
            try:
                params = read_all_params(detail_paths)
                print(f"     VL: {params['vl_min']}÷{params['vl_max']}V | "
                      f"PF: {pf_level(params['pf'])} | "
                      f"ΔU={params['delta_u']}% ΔI={params['delta_i']}% | "
                      f"THD={params['thd_max']}% TDD={params['tdd_max']}%")
                nhanxet = build_nhanxet(ma_so, ten_may, params)
            except Exception:
                print(f"     [OCR thất bại] dùng placeholder")
                traceback.print_exc()
                nhanxet = (
                    f"Biểu đồ dòng điện tiêu thụ tại {expand_tc(ten_may)} biến đổi... "
                    "trong thời gian đo kiểm. Hệ số công suất cosφ ở mức... "
                    "Điện áp dao động từ [Min] ÷ [Max] V, độ lệch chuẩn của điện áp δU (= ...% ÷ ...%) "
                    "đạt tiêu chuẩn (-5,0% ≤ δ ≤ 5,0%). "
                    "Độ lệch pha điện áp và dòng điện đều ở mức thấp (ΔU = ...% < 5,0%, ΔI = ...% < 10,0%). "
                    "Tổng biến dạng sóng hài điện áp và dòng điện đều ở mức cho phép "
                    "(THDmax = ...% < 8,0% & TDDmax = ...% < 20,0%)."
                )

            self.inject(ma_so, ten_may, folder_path, trend_imgs[0], detail_imgs[:6], nhanxet)
            ok += 1

        print(f"\n  Tổng: {ok} máy | {skip} bỏ qua")
        self.save()


# ==============================================================================
if __name__ == "__main__":
    KewReportBuilder().build()
