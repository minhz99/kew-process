let EDITED_FILES = [];
let PENDING_FILES = []; // Array of { id, file, url }
let SLOT_MAPPING = [null, null, null, null, null, null]; // slot index → PENDING_FILES entry

    function handleBulkDrop(e) {
      e.preventDefault();
      handleBulkFiles(e.dataTransfer.files);
    }

    document.getElementById('ei-bulk-upload-area').onclick = () => document.getElementById('ei-bulk-input').click();

    function handleBulkFiles(files) {
      const container = document.getElementById('ei-preview-container');
      Array.from(files).forEach(file => {
        if (!file.name.toLowerCase().endsWith('.bmp')) return;
        const id = 'f_' + Math.random().toString(36).substr(2, 9);
        const url = URL.createObjectURL(file);
        const fileObj = { id, file, url };
        PENDING_FILES.push(fileObj);

        const el = document.createElement('div');
        el.className = 'ei-pending-preview';
        el.id = id;
        el.draggable = true;
        el.ondragstart = (e) => e.dataTransfer.setData('text/plain', id);
        el.style = 'border: 1px solid var(--border); border-radius: 6px; padding: 5px; background: var(--surface); text-align: center; cursor: grab;';
        el.innerHTML = `
        <img src="${url}" style="width: 100%; aspect-ratio: 4/3; object-fit: contain; border-radius: 4px; margin-bottom: 5px;">
        <div style="font-size: 0.65rem; color: var(--text-muted); white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${file.name}</div>
    `;
        container.appendChild(el);
      });
    }

    function handleDropToSlot(e, slotIdx) {
      e.preventDefault();
      const fileId = e.dataTransfer.getData('text/plain');
      const fileObj = PENDING_FILES.find(f => f.id === fileId);
      if (!fileObj) return;

      SLOT_MAPPING[slotIdx] = fileObj;

      const slot = document.querySelectorAll('.ei-slot')[slotIdx];
      const content = slot.querySelector('.ei-slot-content');
      content.innerHTML = `
    <div style="position: relative; width: 100%;">
        <img src="${fileObj.url}" style="width: 100%; aspect-ratio: 4/3; object-fit: contain; border-radius: 4px;">
        <div style="font-size: 0.6rem; margin-top: 2px; color: var(--text); overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${fileObj.file.name}</div>
        <button onclick="clearSlot(${slotIdx}); event.stopPropagation();" style="position: absolute; top: -5px; right: -5px; background: var(--danger); color: white; border: none; border-radius: 50%; width: 15px; height: 15px; font-size: 10px; cursor: pointer;">×</button>
    </div>
  `;
    }

    function clearSlot(idx) {
      SLOT_MAPPING[idx] = null;
      const slot = document.querySelectorAll('.ei-slot')[idx];
      slot.querySelector('.ei-slot-content').innerHTML = '';
    }


    async function submitEditImages() {
      const errorEl = document.getElementById('ei-error');
      const gallery = document.getElementById('ei-gallery');
      const resultsContainer = document.getElementById('ei-results-container');
      const btnSubmit = document.getElementById('btn-edit-img-submit');
      const btnZip = document.getElementById('btn-edit-img-zip');
      const spinner = document.getElementById('ei-spinner');
      const btnText = document.getElementById('ei-btn-text');

      const selectedFiles = [];
      for (let i = 0; i < 6; i++) {
        if (SLOT_MAPPING[i]) {
          selectedFiles.push({ idx: i, file: SLOT_MAPPING[i].file });
        }
      }

      if (selectedFiles.length === 0) {
        errorEl.textContent = 'Vui lòng kéo thả ít nhất một ảnh vào các ô chức năng bên trên.';
        errorEl.style.display = 'block';
        return;
      }

      errorEl.style.display = 'none';
      gallery.innerHTML = '';
      resultsContainer.style.display = 'block';
      btnSubmit.disabled = true;
      btnZip.style.display = 'none';
      spinner.style.display = 'inline-block';
      btnText.textContent = 'Đang xử lý...';
      EDITED_FILES = [];

      // Thu thập thông số đo
      const parameters = {};
      const fields = ['V1', 'V2', 'V3', 'A1', 'A2', 'A3', 'P1', 'P2', 'P3', 'Q1', 'Q2', 'Q3',
                      'S1', 'S2', 'S3', 'PF1', 'PF2', 'PF3', 'Vdeg1', 'Vdeg2', 'Vdeg3',
                      'Adeg1', 'Adeg2', 'Adeg3', 'THDV1', 'THDV2', 'THDV3',
                      'THDA1', 'THDA2', 'THDA3', 'P', 'Q', 'S', 'PF', 'freq', 'An', 'V_unb', 'A_unb'];
      fields.forEach(f => {
        const el = document.getElementById('ei-' + f);
        if (el && el.value && el.value.trim() !== "") {
          parameters[f] = el.value.trim().replace(/,/g, '.');
        }
      });

      // Thu thập thời gian (nếu có điền bất kỳ trường nào)
      const tsDD   = (document.getElementById('ei-ts-dd')  ?.value || '').trim();
      const tsMo   = (document.getElementById('ei-ts-mo')  ?.value || '').trim();
      const tsYYYY = (document.getElementById('ei-ts-yyyy')?.value || '').trim();
      const tsHH   = (document.getElementById('ei-ts-hh')  ?.value || '').trim();
      const tsMi   = (document.getElementById('ei-ts-mi')  ?.value || '').trim();
      const tsSS   = (document.getElementById('ei-ts-ss')  ?.value || '').trim();
      const hasTimestamp = Boolean(tsDD || tsMo || tsYYYY || tsHH || tsMi || tsSS);
      let timestampsForFiles = [];
      try {
        timestampsForFiles = buildTimestampPlan({
          hasTimestamp,
          tsDD,
          tsMo,
          tsYYYY,
          tsHH,
          tsMi,
          tsSS,
          fileCount: selectedFiles.length
        });
      } catch (err) {
        errorEl.textContent = err.message;
        errorEl.style.display = 'block';
        btnSubmit.disabled = false;
        spinner.style.display = 'none';
        btnText.textContent = '📸 Tiếp tục xử lý';
        return;
      }

      for (let order = 0; order < selectedFiles.length; order++) {
        const item = selectedFiles[order];
        const file = item.file;
        const i = item.idx;
        const imageTimestamp = timestampsForFiles[order];
        const card = document.createElement('div');
        card.className = 'chart-card';
        card.style.padding = '10px';
        card.innerHTML = `
      <div style="font-size:0.75rem; color:var(--text-muted); margin-bottom:5px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${file.name}</div>
      <div class="skeleton" style="width:100%; aspect-ratio:4/3; border-radius:6px; background:var(--surface2); animation: pulse 1.5s infinite;"></div>
    `;
        gallery.appendChild(card);
        card.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

        try {
          // Bước 1: xử lý thông số đo
          const templateSelect = document.getElementById('ei-template-select');
          const meterModel = templateSelect ? templateSelect.value : 'kew6315';
          let blob = await processImageServerSide(file, i, parameters, meterModel);

          // Bước 2: áp dụng timestamp (nếu có điền)
          if (imageTimestamp) {
            blob = await applyTimestampServerSide(blob, imageTimestamp);
          }

          const url = URL.createObjectURL(blob);
          const tsLabel = imageTimestamp
            ? `<div style="font-size:0.63rem; color:var(--text-muted); margin-top:3px; text-align:center;">⏱ ${imageTimestamp}</div>`
            : '';

          card.innerHTML = `
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:5px;">
          <div style="font-size:0.75rem; color:var(--text-muted); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width:70%;">${file.name}</div>
          <a href="${url}" download="Edited_${file.name}" style="font-size:0.7rem; color:var(--accent); text-decoration:none;">Lưu ⬇️</a>
        </div>
        <img src="${url}" style="width:100%; border-radius:4px; cursor:pointer;" onclick="window.open('${url}')" />
        ${tsLabel}
      `;
          EDITED_FILES.push({ blob, name: file.name });
        } catch (err) {
          card.innerHTML = `<div style="color:var(--danger); font-size:0.75rem; padding:20px;">Lỗi: ${err.message}</div>`;
        }
      }

      btnSubmit.disabled = false;
      spinner.style.display = 'none';
      btnText.textContent = '📸 Tiếp tục xử lý';
      btnZip.style.display = EDITED_FILES.length > 0 ? 'block' : 'none';
    }

    async function downloadEditedZip() {
      if (EDITED_FILES.length === 0) return;
      const btn = document.getElementById('btn-edit-img-zip');
      const originalText = btn.innerHTML;
      btn.disabled = true;
      btn.textContent = '⏳ Đang nén...';

      const zip = new JSZip();
      EDITED_FILES.forEach(item => {
        zip.file('Edited_' + item.name, item.blob);
      });

      try {
        const content = await zip.generateAsync({ type: 'blob' });
        const url = URL.createObjectURL(content);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'Edited_Meter_Images.zip';
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
      } catch (err) {
        const errorEl = document.getElementById('ei-error');
        errorEl.textContent = `Không thể tạo ZIP: ${err.message}`;
        errorEl.style.display = 'block';
      } finally {
        btn.disabled = false;
        btn.innerHTML = originalText;
      }
    }


// ════════════════════════════════════════════════════════════════════
//  API HELPERS
// ════════════════════════════════════════════════════════════════════

async function processImageServerSide(file, screenIdx, params, meterModel) {
    const formData = new FormData();
    formData.append('file', file);
    formData.append('screenIdx', screenIdx);
    formData.append('parameters', JSON.stringify(params));
    formData.append('meterModel', meterModel);

    const response = await fetch('/api/image/process', {
        method: 'POST',
        body: formData
    });

    if (!response.ok) {
        let errMsg = `Lỗi HTTP ${response.status}`;
        try {
            const errData = await response.json();
            if (errData && errData.error) errMsg = errData.error;
        } catch (e) { /* ignore */ }
        throw new Error(errMsg);
    }

    return await response.blob();
}

/** Gửi blob (đã xử lý thông số) lên server để dán chữ số thời gian */
async function applyTimestampServerSide(blobOrFile, timestamp) {
    const formData = new FormData();
    const fname = blobOrFile instanceof File ? blobOrFile.name : 'edited.bmp';
    formData.append('file', blobOrFile, fname);
    formData.append('timestamp', timestamp);

    const response = await fetch('/api/image/apply-timestamp', {
        method: 'POST',
        body: formData,
    });

    if (!response.ok) {
        let errMsg = `Lỗi HTTP ${response.status} (timestamp)`;
        try {
            const errData = await response.json();
            if (errData && errData.error) errMsg = errData.error;
        } catch (_) {}
        throw new Error(errMsg);
    }
    return await response.blob();
}


// ════════════════════════════════════════════════════════════════════
//  TIMESTAMP UI HELPERS
// ════════════════════════════════════════════════════════════════════

/** Điền thời gian hiện tại vào các ô nhập */
function fillTimestampNow() {
    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    document.getElementById('ei-ts-dd').value   = pad(now.getDate());
    document.getElementById('ei-ts-mo').value   = pad(now.getMonth() + 1);
    document.getElementById('ei-ts-yyyy').value = now.getFullYear();
    document.getElementById('ei-ts-hh').value   = pad(now.getHours());
    document.getElementById('ei-ts-mi').value   = pad(now.getMinutes());
    document.getElementById('ei-ts-ss').value   = pad(now.getSeconds());
    // Cập nhật picker
    const pv = `${now.getFullYear()}-${pad(now.getMonth()+1)}-${pad(now.getDate())}T${pad(now.getHours())}:${pad(now.getMinutes())}`;
    document.getElementById('ei-ts-picker').value = pv;
}

/** Đồng bộ từ datetime-local picker vào các ô riêng */
function fillTimestampFromPicker(val) {
    if (!val) return;
    const [datePart, timePart] = val.split('T');
    const [yyyy, mo, dd] = datePart.split('-');
    const [hh, mi]       = timePart.split(':');
    document.getElementById('ei-ts-dd').value   = dd;
    document.getElementById('ei-ts-mo').value   = mo;
    document.getElementById('ei-ts-yyyy').value = yyyy;
    document.getElementById('ei-ts-hh').value   = hh;
    document.getElementById('ei-ts-mi').value   = mi || '00';
}

function buildTimestampPlan(opts) {
    const { hasTimestamp, tsDD, tsMo, tsYYYY, tsHH, tsMi, tsSS, fileCount } = opts;
    if (!hasTimestamp) {
        return Array(fileCount).fill(null);
    }

    const baseTimestamp = formatTimestampString(tsDD, tsMo, tsYYYY, tsHH, tsMi, tsSS);
    if (fileCount <= 1) {
        return [baseTimestamp];
    }

    const minStep = parseNonNegativeInt(
        document.getElementById('ei-ts-step-min')?.value || '',
        'm (giây)',
        false
    );
    const maxStep = parseNonNegativeInt(
        document.getElementById('ei-ts-step-max')?.value || '',
        'n (giây)',
        false
    );
    if (maxStep < minStep) {
        throw new Error('Khoảng ngẫu nhiên không hợp lệ: n phải lớn hơn hoặc bằng m.');
    }

    if (minStep === 0 && maxStep === 0) {
        return Array(fileCount).fill(baseTimestamp);
    }

    const baseDate = parseStrictTimestampDate(tsDD, tsMo, tsYYYY, tsHH, tsMi, tsSS);
    if (!baseDate) {
        throw new Error('Để cộng thời gian giữa các ảnh, vui lòng nhập đầy đủ ngày/tháng/năm giờ:phút:giây hợp lệ.');
    }

    const planned = [baseTimestamp];
    let current = new Date(baseDate.getTime());
    for (let i = 1; i < fileCount; i++) {
        const delta = randomIntInclusive(minStep, maxStep);
        current = new Date(current.getTime() + (delta * 1000));
        planned.push(formatTimestampFromDate(current));
    }

    return planned;
}

function parseNonNegativeInt(rawValue, label, allowEmpty) {
    const raw = String(rawValue || '').trim();
    if (raw === '') {
        if (allowEmpty) return 0;
        throw new Error(`Vui lòng nhập ${label}.`);
    }
    if (!/^\d+$/.test(raw)) {
        throw new Error(`${label} phải là số nguyên không âm.`);
    }
    return parseInt(raw, 10);
}

function parseStrictTimestampDate(dd, mo, yyyy, hh, mi, ss) {
    const parts = [dd, mo, yyyy, hh, mi, ss].map(v => String(v || '').trim());
    if (!parts.every(v => /^\d+$/.test(v))) {
        return null;
    }

    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10);
    const year = parseInt(parts[2], 10);
    const hour = parseInt(parts[3], 10);
    const minute = parseInt(parts[4], 10);
    const second = parseInt(parts[5], 10);

    if (year < 1000 || month < 1 || month > 12 || day < 1 || day > 31) return null;
    if (hour < 0 || hour > 23 || minute < 0 || minute > 59 || second < 0 || second > 59) return null;

    const d = new Date(year, month - 1, day, hour, minute, second, 0);
    if (d.getFullYear() !== year || d.getMonth() !== month - 1 || d.getDate() !== day) return null;
    return d;
}

function formatTimestampString(dd, mo, yyyy, hh, mi, ss) {
    return `${String(dd || '00').padStart(2, '0')}/${String(mo || '00').padStart(2, '0')}/${String(yyyy || '0000').padStart(4, '0')} ${String(hh || '00').padStart(2, '0')}:${String(mi || '00').padStart(2, '0')}:${String(ss || '00').padStart(2, '0')}`;
}

function formatTimestampFromDate(d) {
    const pad = n => String(n).padStart(2, '0');
    return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

function randomIntInclusive(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}
