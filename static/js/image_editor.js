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

        // Create preview element
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

      // Update mapping
      SLOT_MAPPING[slotIdx] = fileObj;

      // Visual update
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
      spinner.style.display = 'inline-block';
      btnText.textContent = 'Đang xử lý...';
      EDITED_FILES = [];

      const parameters = {};
      const fields = ['V1', 'V2', 'V3', 'A1', 'A2', 'A3', 'P1', 'P2', 'P3', 'Q1', 'Q2', 'Q3', 'S1', 'S2', 'S3', 'PF1', 'PF2', 'PF3', 'Vdeg1', 'Vdeg2', 'Vdeg3', 'Adeg1', 'Adeg2', 'Adeg3', 'THDV1', 'THDV2', 'THDV3', 'THDA1', 'THDA2', 'THDA3', 'P', 'Q', 'S', 'PF', 'freq', 'An', 'V_unb', 'A_unb'];

      fields.forEach(f => {
        const el = document.getElementById('ei-' + f);
        if (el && el.value && el.value.trim() !== "") {
          // Chuẩn hoá: dấu phẩy thập phân -> dấu chấm
          parameters[f] = el.value.trim().replace(/,/g, '.');
        }
      });

      for (let item of selectedFiles) {
        const file = item.file;
        const i = item.idx;
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
          // Server processing
          const templateSelect = document.getElementById('ei-template-select');
          const meterModel = templateSelect ? templateSelect.value : 'kew6315';
          const blob = await processImageServerSide(file, i, parameters, meterModel);
          const url = URL.createObjectURL(blob);


          card.innerHTML = `
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:5px;">
          <div style="font-size:0.75rem; color:var(--text-muted); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; max-width:70%;">${file.name}</div>
          <a href="${url}" download="Edited_${file.name}" style="font-size:0.7rem; color:var(--accent); text-decoration:none;">Lưu ⬇️</a>
        </div>
        <img src="${url}" style="width:100%; border-radius:4px; cursor:pointer;" onclick="window.open('${url}')" />
      `;
          EDITED_FILES.push({ blob, name: file.name });
        } catch (err) {
          card.innerHTML = `<div style="color:var(--danger); font-size:0.75rem; padding:20px;">Lỗi: ${err.message}</div>`;
        }
      }

      btnSubmit.disabled = false;
      spinner.style.display = 'none';
      btnText.textContent = '📸 Tiếp tục xử lý';
      if (EDITED_FILES.length > 0) btnZip.style.display = 'block';
    }

    function downloadEditedZip() {
      if (EDITED_FILES.length === 0) return;
      const btn = document.getElementById('btn-edit-img-zip');
      const originalText = btn.innerHTML;
      btn.disabled = true;
      btn.textContent = '⏳ Đang nén...';

      const zip = new JSZip();
      EDITED_FILES.forEach(item => {
        zip.file('Edited_' + item.name, item.blob);
      });

      zip.generateAsync({ type: "blob" }).then(function (content) {
        const url = URL.createObjectURL(content);
        const a = document.createElement('a'); a.href = url; a.download = 'Edited_Meter_Images.zip';
        document.body.appendChild(a); a.click(); a.remove();
        btn.disabled = false; btn.innerHTML = originalText;
      });
    }



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
        } catch (e) {
            // ignore
        }
        throw new Error(errMsg);
    }

    return await response.blob();
}
