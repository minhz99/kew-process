import re

with open('dashboard.html', 'r', encoding='utf-8') as f:
    text = f.read()

# 1. Add jszip
text = text.replace('</head>', '<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>\n</head>')

# 2. Replacing UI File Inputs
old_ui = """  <div class="chart-card" style="margin-bottom: 1.5rem;">
    <h3>Chọn các tệp ảnh đo (.BMP)</h3>
    <input type="file" id="edit-img-file" multiple accept=".bmp" style="margin-bottom: 1rem; display: block; padding: 10px; border: 1px dashed var(--border); border-radius: 8px; width: 100%; color: var(--text); background: var(--surface2);"/>
    <div style="font-size: 0.85rem; color: var(--text-muted); line-height: 1.6;">
      Chọn một hoặc nhiều tệp ảnh gốc. Hệ thống sẽ xử lý và hiển thị kết quả trực tiếp bên dưới. 
      Bạn có thể tải về từng ảnh hoặc toàn bộ dưới dạng file nén.
    </div>
  </div>"""

new_ui = """  <div class="chart-card" style="margin-bottom: 1.5rem;">
    <h3>Tải lên các tệp ảnh (.BMP)</h3>
    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px; margin-bottom: 1rem;">
      <div><label style="font-size:0.8rem; color:var(--text-muted)">Ảnh 1 (Tổng hợp)</label><input type="file" id="ei-file-0" accept=".bmp" style="width:100%; font-size:0.8rem; color:white; background:var(--surface2); border:1px solid var(--border); padding:5px; border-radius:4px;"></div>
      <div><label style="font-size:0.8rem; color:var(--text-muted)">Ảnh 2 (Pha áp)</label><input type="file" id="ei-file-1" accept=".bmp" style="width:100%; font-size:0.8rem; color:white; background:var(--surface2); border:1px solid var(--border); padding:5px; border-radius:4px;"></div>
      <div><label style="font-size:0.8rem; color:var(--text-muted)">Ảnh 3 (Sin áp)</label><input type="file" id="ei-file-2" accept=".bmp" style="width:100%; font-size:0.8rem; color:white; background:var(--surface2); border:1px solid var(--border); padding:5px; border-radius:4px;"></div>
      <div><label style="font-size:0.8rem; color:var(--text-muted)">Ảnh 4 (Sin dòng)</label><input type="file" id="ei-file-3" accept=".bmp" style="width:100%; font-size:0.8rem; color:white; background:var(--surface2); border:1px solid var(--border); padding:5px; border-radius:4px;"></div>
      <div><label style="font-size:0.8rem; color:var(--text-muted)">Ảnh 5 (Hài áp)</label><input type="file" id="ei-file-4" accept=".bmp" style="width:100%; font-size:0.8rem; color:white; background:var(--surface2); border:1px solid var(--border); padding:5px; border-radius:4px;"></div>
      <div><label style="font-size:0.8rem; color:var(--text-muted)">Ảnh 6 (Hài dòng)</label><input type="file" id="ei-file-5" accept=".bmp" style="width:100%; font-size:0.8rem; color:white; background:var(--surface2); border:1px solid var(--border); padding:5px; border-radius:4px;"></div>
    </div>
  </div>"""
text = text.replace(old_ui, new_ui)

# 3. Replacing parameters
# We replace from <div class="charts-grid"> to the checkbox container.
params_pattern = r'<div class="charts-grid">\s*<div class="chart-card">\s*<h3>Thông số cấu hình \(Cơ bản\).*?<!-- this regex will replace up to the fluctuate logic -->'
# Let's be safer:
param_re = r'<div class="charts-grid">\s*<div class="chart-card">\s*<h3>Thông số cấu hình \(Cơ bản\).*?</div>\s*</div>\s*<div class="charts-grid">\s*<div class="chart-card">\s*<h3>Thời gian & THD.*?</div>\s*</div>'

new_params = """<div class="chart-card" style="margin-bottom: 1.5rem;">
    <h3>Thông số cấu hình (38 giá trị)</h3>
    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(130px, 1fr)); gap: 10px; margin-bottom: 15px;">
"""
cols = ['V1','V2','V3','A1','A2','A3','P1','P2','P3','Q1','Q2','Q3','S1','S2','S3','PF1','PF2','PF3','Vdeg1','Vdeg2','Vdeg3','Adeg1','Adeg2','Adeg3','THDV1','THDV2','THDV3','THDA1','THDA2','THDA3','P','Q','S','PF','freq','An','V_unb','A_unb']
for c in cols:
    new_params += f"""      <div><label style="font-size:0.8rem; color:var(--text-muted)">{c}</label><input type="text" id="ei-{c}" placeholder="{c}" style="width:100%; padding:6px; background:var(--surface2); border:1px solid var(--border); color:white; border-radius:4px;"/></div>\n"""
new_params += """    </div>\n  </div>\n<div class="charts-grid">\n"""

text = re.sub(param_re, new_params, text, flags=re.DOTALL)

# 4. JS replacement
js_old = """async function submitEditImages() {
  const fileInput = document.getElementById('edit-img-file');
  const errorEl = document.getElementById('ei-error');
  const gallery = document.getElementById('ei-gallery');
  const resultsContainer = document.getElementById('ei-results-container');
  const btnSubmit = document.getElementById('btn-edit-img-submit');
  const btnZip = document.getElementById('btn-edit-img-zip');
  const spinner = document.getElementById('ei-spinner');
  const btnText = document.getElementById('ei-btn-text');

  if (!fileInput.files.length) {
    errorEl.textContent = 'Vui lòng chọn ít nhất một file ảnh .BMP';
    errorEl.style.display = 'block';
    return;
  }
  
  errorEl.style.display = 'none';
  gallery.innerHTML = '';
  resultsContainer.style.display = 'block';
  btnSubmit.disabled = true;
  spinner.style.display = 'inline-block';
  btnText.textContent = 'Đang xử lý...';
  EDITED_FILES = [];

  // Sort files by name to match screen mapping (SD140, SD141, etc.)
  const files = Array.from(fileInput.files).sort((a, b) => a.name.localeCompare(b.name, undefined, {numeric: true}));
  
  const parameters = {};
  const fields = ['v1', 'v2', 'v3', 'a1', 'p', 'q', 's', 'pf1', 'freq', 'thdv', 'date', 'time'];
  fields.forEach(f => {
    const el = document.getElementById('ei-' + f.toUpperCase()) || document.getElementById('ei-' + f.toLowerCase());
    if (el) parameters[f] = el.value;
  });
  parameters['fluctuate'] = document.getElementById('ei-fluctuate').checked;

  for (let i = 0; i < files.length; i++) {
    const file = files[i];"""

js_submit = """async function submitEditImages() {
  const errorEl = document.getElementById('ei-error');
  const gallery = document.getElementById('ei-gallery');
  const resultsContainer = document.getElementById('ei-results-container');
  const btnSubmit = document.getElementById('btn-edit-img-submit');
  const btnZip = document.getElementById('btn-edit-img-zip');
  const spinner = document.getElementById('ei-spinner');
  const btnText = document.getElementById('ei-btn-text');

  const selectedFiles = [];
  for (let i = 0; i < 6; i++) {
    const p = document.getElementById('ei-file-' + i);
    if (p && p.files.length > 0) {
      selectedFiles.push({ idx: i, file: p.files[0] });
    }
  }

  if (selectedFiles.length === 0) {
    errorEl.textContent = 'Vui lòng chọn ít nhất một file ảnh .BMP tại bất kỳ ô nào.';
    errorEl.style.display = 'block';
    return;
  }
  
  errorEl.style.display = 'none';
  gallery.innerHTML = '';
  resultsContainer.style.display = 'block';
  btnSubmit.disabled = true;
  spinner.style.display = 'inline-block';
  btnText.textContent = 'Đang xử lý...';
  EDITED_FILES = [];
  
  const parameters = {};
  const fields = ['V1','V2','V3','A1','A2','A3','P1','P2','P3','Q1','Q2','Q3','S1','S2','S3','PF1','PF2','PF3','Vdeg1','Vdeg2','Vdeg3','Adeg1','Adeg2','Adeg3','THDV1','THDV2','THDV3','THDA1','THDA2','THDA3','P','Q','S','PF','freq','An','V_unb','A_unb'];
  
  fields.forEach(f => {
    const el = document.getElementById('ei-' + f);
    if (el && el.value && el.value.trim() !== "") {
      parameters[f] = el.value.trim();
    }
  });
  parameters['fluctuate'] = document.getElementById('ei-fluctuate').checked;

  for (let item of selectedFiles) {
    const file = item.file;
    const i = item.idx;"""

text = text.replace(js_old, js_submit)

# Also fix the form append to use `idx` instead of loop counter `i` for screen. Wait, I mapped `i` to `item.idx` above! 
# Let's ensure the var names align:
js_form_1 = """fd.append('idx', i); // screen index"""
js_form_2 = """fd.append('idx', i); // screen index"""
# Above logic holds because I declared `const i = item.idx;`

# 5. Fix download zip
old_dl = """function downloadEditedZip() {
  const fileInput = document.getElementById('edit-img-file');
  if (!fileInput.files.length) return;"""

old_dl_full = re.search(r'function downloadEditedZip\(\) \{.*?\n\}\n', text, re.DOTALL).group(0)

new_dl = """function downloadEditedZip() {
  if (EDITED_FILES.length === 0) return;
  const btn = document.getElementById('btn-edit-img-zip');
  const originalText = btn.innerHTML;
  btn.disabled = true;
  btn.textContent = '⏳ Đang nén...';
  
  const zip = new JSZip();
  EDITED_FILES.forEach(item => {
    zip.file('Edited_' + item.name, item.blob);
  });
  
  zip.generateAsync({type:"blob"}).then(function(content) {
    const url = URL.createObjectURL(content);
    const a = document.createElement('a'); a.href = url; a.download = 'Edited_Meter_Images.zip';
    document.body.appendChild(a); a.click(); a.remove();
    btn.disabled = false; btn.innerHTML = originalText;
  });
}
"""
text = text.replace(old_dl_full, new_dl)

with open('dashboard.html', 'w', encoding='utf-8') as f:
    f.write(text)
print("PATCH_SUCCESS")
