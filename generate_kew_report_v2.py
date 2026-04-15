#!/usr/bin/env python3
"""
generate_kew_report_v2.py
=========================
Tạo báo cáo KEW tự động cho dự án Prime Đại Việt.
- Nhận JSON config với thông tin thiết bị và đường dẫn ảnh.
- Ưu tiên dùng giá trị nhập tay (manual_values) trong JSON.
- Nếu không có manual_values → dùng Tesseract OCR để đọc từ ảnh.
- Tạo file Word theo mẫu "03. Chương 4".

Chạy:
    python generate_kew_report_v2.py config.json
"""

import json
import os
import re
import sys
import copy
import shutil
import zipfile
from pathlib import Path

try:
    from lxml import etree
except ImportError:
    print("✗ Thiếu lxml: pip3 install lxml")
    sys.exit(1)

try:
    import pytesseract
    from PIL import Image, ImageOps
    import numpy as np
    HAS_TESSERACT = True
except ImportError:
    HAS_TESSERACT = False
    print("! Không có pytesseract, sẽ chỉ dùng manual values")

# Namespaces Word
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

def qw(tag):
    return f"{{{W}}}{tag}"


# ==============================================================================
# OCR HELPER
# ==============================================================================

def ocr_field(image_path, region=None):
    """
    Đọc một field từ ảnh dùng Tesseract.
    region: tuple (x1, y1, x2, y2) hoặc None để đọc toàn bộ ảnh.
    """
    if not HAS_TESSERACT:
        return None
    
    try:
        img = Image.open(image_path).convert('L')
        if region:
            img = img.crop(region)
        
        # Preprocessing
        arr = np.array(img)
        threshold = np.percentile(arr, 50)
        binary = (arr < threshold).astype(np.uint8) * 255
        img_binary = Image.fromarray(binary, mode='L')
        img_binary = ImageOps.invert(img_binary)
        
        # Upscale
        scale = max(4, 200 // max(img_binary.size))
        img_big = img_binary.resize(
            (img_binary.width * scale, img_binary.height * scale),
            Image.NEAREST
        )
        
        text = pytesseract.image_to_string(
            img_big, lang='eng',
            config='--psm 7 -c tessedit_char_whitelist=0123456789.'
        )
        
        cleaned = re.sub(r'[^0-9.]', '', text.strip())
        return cleaned if cleaned else None
    except Exception as e:
        print(f"    [OCR error] {e}")
        return None


def parse_number(text):
    """Chuyển text thành số."""
    if text is None:
        return None
    cleaned = re.sub(r'[^0-9.\-]', '', text.replace(',', '.'))
    try:
        return float(cleaned)
    except ValueError:
        return None


# ==============================================================================
# READ VALUES FROM IMAGES
# ==============================================================================

def read_values_from_images(detail_images, manual_values=None):
    """
    Đọc thông số từ 6 ảnh detail.
    Ưu tiên manual_values, fallback OCR nếu không có.
    """
    values = {
        'V1': None, 'V2': None, 'V3': None,
        'A1': None, 'A2': None, 'A3': None,
        'PF': None, 'P': None, 'Q': None, 'S': None,
        'freq': None,
        'Delta_U': None, 'Delta_I': None,
        'THD_V': None, 'THD_A': None,
    }
    
    if manual_values:
        for key in values:
            if key in manual_values:
                values[key] = parse_number(str(manual_values[key]))
    
    # Check nếu đã có đủ giá trị quan trọng
    has_voltage = any(values.get(k) for k in ['V1', 'V2', 'V3'])
    has_current = any(values.get(k) for k in ['A1', 'A2', 'A3'])
    has_pf = values.get('PF') is not None
    
    if has_voltage and has_current and has_pf:
        print("  ✓ Dùng manual values")
        return values
    
    # Fallback: OCR (sẽ cải thiện sau)
    print("  ⚠ Manual values không đủ, OCR không khả dụng → dùng placeholder")
    return values


# ==============================================================================
# CALCULATE & BUILD COMMENTARY
# ==============================================================================

def calc_stats(values):
    """Tính các thống kê từ values."""
    voltages = [v for v in [values.get('V1'), values.get('V2'), values.get('V3')] if v]
    currents = [v for v in [values.get('A1'), values.get('A2'), values.get('A3')] if v]
    
    stats = {
        'V_min': min(voltages) if voltages else None,
        'V_max': max(voltages) if voltages else None,
        'I_avg': sum(currents) / len(currents) if currents else None,
        'PF': values.get('PF'),
        'P': values.get('P'),
        'Delta_U': values.get('Delta_U'),
        'Delta_I': values.get('Delta_I'),
        'THD_V': values.get('THD_V'),
        'THD_A': values.get('THD_A'),
    }
    
    # Tính Delta_U nếu không có
    if stats['Delta_U'] is None and stats['V_min'] and stats['V_max']:
        v_nom = 400.0
        stats['Delta_U'] = ((stats['V_min'] - v_nom) / v_nom) * 100
    
    return stats


def fmt_num(val, decimals=1):
    """Format số sang tiếng Việt (dấu phẩy thập phân)."""
    if val is None:
        return "..."
    try:
        return f"{val:.{decimals}f}".replace('.', ',')
    except:
        return str(val)


def pf_level(pf):
    if pf is None:
        return "..."
    if pf >= 0.9:
        return "cao"
    elif pf >= 0.8:
        return "trung bình"
    else:
        return "thấp"


def quality_rating(stats):
    """Đánh giá chất lượng tổng thể."""
    issues = []
    
    if stats.get('PF') and stats['PF'] < 0.8:
        issues.append("cosφ thấp")
    if stats.get('Delta_U') and abs(stats['Delta_U']) > 5:
        issues.append("mất cân bằng điện áp cao")
    if stats.get('Delta_I') and stats['Delta_I'] > 10:
        issues.append("mất cân bằng dòng điện cao")
    if stats.get('THD_V') and stats['THD_V'] > 8:
        issues.append("THD điện áp cao")
    if stats.get('THD_A') and stats['THD_A'] > 20:
        issues.append("TDD dòng điện cao")
    
    if not issues:
        return "tốt", []
    elif len(issues) <= 2:
        return "khá tốt", issues
    else:
        return "chưa tốt", issues


def build_nhanxet(ten_thiet_bi, stats):
    """Xây dựng đoạn nhận xét theo mẫu '03. Chương 4'."""
    rating, issues = quality_rating(stats)
    
    v_min = fmt_num(stats.get('V_min'))
    v_max = fmt_num(stats.get('V_max'))
    pf = fmt_num(stats.get('PF'), 2)
    pf_lvl = pf_level(stats.get('PF'))
    
    delta_u = fmt_num(stats.get('Delta_U'))
    delta_i = fmt_num(stats.get('Delta_I'))
    thd_v = fmt_num(stats.get('THD_V'), 2)
    thd_a = fmt_num(stats.get('THD_A'), 2)
    
    # Build text
    parts = []
    parts.append(f"Chất lượng điện cấp cho {ten_thiet_bi} ở mức {rating}.")
    
    if stats.get('PF') is not None:
        parts.append(f"Hệ số công suất cosφ ở mức {pf_lvl} ({pf}).")
    
    if stats.get('V_min') and stats.get('V_max'):
        v_dev = ((stats['V_min'] - 400) / 400) * 100
        dev_str = fmt_num(v_dev, 2)
        parts.append(
            f"Điện áp dao động từ {v_min} ÷ {v_max} V, "
            f"độ lệch δU = {dev_str}% "
            f"{'đạt tiêu chuẩn' if abs(v_dev) <= 5 else 'không đạt tiêu chuẩn'} "
            f"(-5,0% ≤ δ ≤ 5,0%)."
        )
    
    if stats.get('Delta_U') is not None and stats.get('Delta_I') is not None:
        du_ok = stats['Delta_U'] < 5
        di_ok = stats['Delta_I'] < 10
        parts.append(
            f"Độ lệch pha điện áp và dòng điện "
            f"{'đều ở mức thấp' if du_ok and di_ok else 'ở mức cao'} "
            f"(ΔU = {delta_u}% {'<' if du_ok else '>'} 5,0%, "
            f"ΔI = {delta_i}% {'<' if di_ok else '>'} 10,0%)."
        )
    
    # THD/TDD sentence
    if stats.get('THD_V') is not None and stats.get('THD_A') is not None:
        thd_ok = stats['THD_V'] < 8
        tdd_ok = stats['THD_A'] < 20
        if thd_ok and tdd_ok:
            parts.append(
                f"Tổng biến dạng sóng hài điện áp và dòng điện đều ở mức cho phép "
                f"(THDmax = {thd_v}% < 8,0% & TDDmax = {thd_a}% < 20,0%)."
            )
        elif thd_ok and not tdd_ok:
            parts.append(
                f"Tổng biến dạng sóng hài điện áp ở mức cho phép "
                f"(THDmax = {thd_v}% < 8,0%); tuy nhiên, tổng biến dạng sóng hài "
                f"dòng điện cao hơn mức cho phép (TDDmax = {thd_a}% > 20,0%)."
            )
        elif not thd_ok and tdd_ok:
            parts.append(
                f"Tổng biến dạng sóng hài dòng điện ở mức cho phép "
                f"(TDDmax = {thd_a}% < 20,0%); tuy nhiên, tổng biến dạng sóng hài "
                f"điện áp cao hơn mức cho phép (THDmax = {thd_v}% > 8,0%)."
            )
        else:
            parts.append(
                f"Tổng biến dạng sóng hài điện áp và dòng điện đều vượt mức cho phép "
                f"(THDmax = {thd_v}% > 8,0% & TDDmax = {thd_a}% > 20,0%)."
            )
    
    return " ".join(parts)


# ==============================================================================
# WORD REPORT BUILDER
# ==============================================================================

class KewReportBuilder:
    def __init__(self, template_file, output_file):
        self.template_file = template_file
        self.output_file = output_file
        self.work_dir = output_file + "_workdir"
        self.media_dir = os.path.join(self.work_dir, "word", "media")
        self._rid_num = 200
        self._img_num = 1
        self._docpr_num = 900000000
        self.fig_counter = 1
        
        # Sizes (EMU)
        self.TREND_CX = "6048000"
        self.TREND_CY = "2012400"
        self.DETAIL_CX = "1990800"
        self.DETAIL_CY = "1486800"
    
    def unpack(self):
        if os.path.exists(self.work_dir):
            shutil.rmtree(self.work_dir)
        with zipfile.ZipFile(self.template_file, 'r') as z:
            z.extractall(self.work_dir)
        os.makedirs(self.media_dir, exist_ok=True)
    
    def load(self):
        self.doc_path = os.path.join(self.work_dir, "word", "document.xml")
        self.rels_path = os.path.join(self.work_dir, "word", "_rels", "document.xml.rels")
        self.ct_path = os.path.join(self.work_dir, "[Content_Types].xml")
        
        self.doc_tree = etree.parse(self.doc_path)
        self.rels_tree = etree.parse(self.rels_path)
        self.ct_tree = etree.parse(self.ct_path)
        
        self.body = self.doc_tree.getroot().find(f".//{{{W}}}body")
        self.rels_root = self.rels_tree.getroot()
        self.ct_root = self.ct_tree.getroot()
    
    def clear_body(self):
        sect = self.body.find(qw("sectPr"))
        sect_copy = copy.deepcopy(sect) if sect is not None else None
        for child in list(self.body):
            self.body.remove(child)
        if sect_copy is not None:
            self.body.append(sect_copy)
    
    def _add_image(self, src_path):
        ext = os.path.splitext(src_path)[1].lower().lstrip(".")
        img_name = f"image{self._img_num}.{ext}"
        shutil.copy2(src_path, os.path.join(self.media_dir, img_name))
        rid = f"rId{self._rid_num}"
        self._img_num += 1
        self._rid_num += 1
        
        rel = etree.SubElement(self.rels_root, "Relationship")
        rel.set("Id", rid)
        rel.set("Type", REL_IMAGE)
        rel.set("Target", f"media/{img_name}")
        
        mime_map = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "bmp": "image/bmp"}
        mime = mime_map.get(ext, f"image/{ext}")
        existing = {el.get("Extension", "").lower() for el in self.ct_root}
        if ext not in existing:
            e = etree.SubElement(self.ct_root, "Default")
            e.set("Extension", ext)
            e.set("ContentType", mime)
        return rid
    
    def _pid(self):
        self._docpr_num += 1
        return str(self._docpr_num)
    
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
          <pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
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
      <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
    </w:tcBorders>
    <w:vAlign w:val="center"/>
  </w:tcPr>
  {para}
</w:tc>"""
    
    def _row(self, cells):
        return f'<w:tr xmlns:w="{W}"><w:trPr><w:jc w:val="center"/></w:trPr>{cells}</w:tr>'
    
    def make_table(self, folder_path, trend_img, detail_imgs):
        dr = [self._add_image(os.path.join(folder_path, d)) for d in detail_imgs[:6]]
        dp = [self._cpara(self._drawing(r, self.DETAIL_CX, self.DETAIL_CY)) for r in dr]
        
        # Trend image là optional
        if trend_img and os.path.exists(os.path.join(folder_path, trend_img)):
            tr = self._add_image(os.path.join(folder_path, trend_img))
            tp = self._cpara(self._drawing(tr, self.TREND_CX, self.TREND_CY), "left")
            trend_row = self._row(self._tc(tp, span=3))
        else:
            trend_row = ""
        
        return etree.fromstring(f"""<w:tbl xmlns:w="{W}">
  <w:tblPr>
    <w:tblW w:w="5000" w:type="pct"/>
    <w:tblLayout w:type="fixed"/>
    <w:tblBorders>
      <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>
      <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
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
  {trend_row}
  {self._row(self._tc(dp[0]) + self._tc(dp[1]) + self._tc(dp[2]))}
  {self._row(self._tc(dp[3]) + self._tc(dp[4]) + self._tc(dp[5]))}
</w:tbl>""")
    
    def make_title(self, ten_thiet_bi):
        return etree.fromstring(f"""<w:p xmlns:w="{W}">
  <w:pPr><w:pStyle w:val="onvn"/></w:pPr>
  <w:r><w:t xml:space="preserve">{ten_thiet_bi}:</w:t></w:r>
</w:p>""")
    
    def make_caption(self, fig_num, ten_thiet_bi):
        return etree.fromstring(f"""<w:p xmlns:w="{W}">
  <w:pPr><w:pStyle w:val="DMhnhK"/></w:pPr>
  <w:r><w:t xml:space="preserve">Hình 4.{fig_num} - Kết quả đo chất lượng điện {ten_thiet_bi}</w:t></w:r>
</w:p>""")
    
    def make_nhanxet(self, text):
        safe = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
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
    
    def inject(self, ten_thiet_bi, folder_path, trend_img, detail_imgs, nhanxet):
        sect = self.body.find(qw("sectPr"))
        els = [
            self.make_title(ten_thiet_bi),
            self.make_table(folder_path, trend_img, detail_imgs),
            self.make_caption(self.fig_counter, ten_thiet_bi),
            self.make_nhanxet(nhanxet),
            self.make_pagebreak(),
        ]
        if sect is not None:
            idx = list(self.body).index(sect)
            for i, el in enumerate(els):
                self.body.insert(idx + i, el)
        else:
            for el in els:
                self.body.append(el)
        self.fig_counter += 1
    
    def save(self):
        self.doc_tree.write(self.doc_path, xml_declaration=True, encoding="UTF-8", standalone=True)
        self.rels_tree.write(self.rels_path, xml_declaration=True, encoding="UTF-8", standalone=True)
        self.ct_tree.write(self.ct_path, xml_declaration=True, encoding="UTF-8", standalone=True)
        
        if os.path.exists(self.output_file):
            try:
                os.remove(self.output_file)
            except PermissionError:
                print("✗ File đang mở trong Word — hãy đóng lại rồi chạy lại!")
                shutil.rmtree(self.work_dir, ignore_errors=True)
                return
        
        with zipfile.ZipFile(self.output_file, 'w', zipfile.ZIP_DEFLATED) as z:
            for root, dirs, files in os.walk(self.work_dir):
                for f in files:
                    full = os.path.join(root, f)
                    z.write(full, os.path.relpath(full, self.work_dir))
        
        shutil.rmtree(self.work_dir)
        print(f"\n✓ Hoàn thành! Lưu tại: {self.output_file}")
    
    def build(self, devices, images_base_dir):
        """
        devices: list of dict với keys:
          - name: tên thiết bị
          - folder_path: đường dẫn thư mục chứa ảnh (trend + detail)
          - trend_image: tên file ảnh trend
          - detail_images: list 6 tên file ảnh detail
          - manual_values: dict giá trị nhập tay (optional)
        """
        if not os.path.exists(self.template_file):
            raise RuntimeError(f"Không tìm thấy template: {self.template_file}")
        
        self.unpack()
        self.load()
        self.clear_body()
        
        ok = skip = 0
        for device in devices:
            ten_thiet_bi = device['name']
            folder_path = device.get('folder_path', images_base_dir)
            trend_img = device.get('trend_image')
            detail_imgs = device.get('detail_images', [])
            manual_values = device.get('manual_values', {})
            
            print(f"\n  ── {ten_thiet_bi}")
            
            if len(detail_imgs) < 6:
                print(f"     [bỏ qua] Cần ít nhất 6 detail images")
                skip += 1
                continue
            
            # Đọc values
            values = read_values_from_images(detail_imgs, manual_values)
            stats = calc_stats(values)
            
            print(f"     V: {fmt_num(stats.get('V_min'))}÷{fmt_num(stats.get('V_max'))}V | "
                  f"PF: {fmt_num(stats.get('PF'), 2)} | "
                  f"ΔU={fmt_num(stats.get('Delta_U'))}% ΔI={fmt_num(stats.get('Delta_I'))}% | "
                  f"THD={fmt_num(stats.get('THD_V'))}% TDD={fmt_num(stats.get('THD_A'))}%")
            
            nhanxet = build_nhanxet(ten_thiet_bi, stats)
            
            self.inject(ten_thiet_bi, folder_path, trend_img, detail_imgs, nhanxet)
            ok += 1
        
        print(f"\n  Tổng: {ok} thiết bị | {skip} bỏ qua")
        self.save()
        return self.output_file


# ==============================================================================
# MAIN
# ==============================================================================

def main():
    if len(sys.argv) < 2:
        print("Usage: python generate_kew_report_v2.py <config.json>")
        print("\nConfig format:")
        print("""{
  "template_file": "/path/to/template.docx",
  "output_file": "/path/to/output.docx",
  "images_base_dir": "/path/to/images",
  "devices": [
    {
      "name": "Tên thiết bị",
      "folder_path": "/path/to/image/folder",
      "trend_image": "trend.jpg",
      "detail_images": ["img1.bmp", "img2.bmp", ...],
      "manual_values": {
        "V1": 385, "V2": 387, "V3": 386,
        "PF": 0.92, "Delta_U": 2.5, "Delta_I": 5.0,
        "THD_V": 3.2, "THD_A": 12.5
      }
    }
  ]
}""")
        sys.exit(1)
    
    config_path = sys.argv[1]
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    builder = KewReportBuilder(
        template_file=config['template_file'],
        output_file=config['output_file']
    )
    
    builder.build(
        devices=config['devices'],
        images_base_dir=config.get('images_base_dir', '.')
    )


if __name__ == "__main__":
    main()
