let EDITED_FILES = [];

    let PENDING_FILES = []; // Array of { id, file, url }
    let SLOT_MAPPING = [null, null, null, null, null, null]; // Stores the PENDING_FILES id for each slot

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
          parameters[f] = el.value.trim();
        }
      });
      parameters['fluctuate'] = false;

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
          // Local processing
          if (Object.keys(DIGIT_TEMPLATES).length === 0) await loadDigitTemplates();

          const blob = await processImageClientSide(file, i, parameters);
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



    const SCREENS = [
      {
        id: "SD140",
        overlays: [
          ...makeGrid(["V1", "V2", "V3"], [94, 158, 222], 54, "w"),
          ...makeGrid(["A1", "A2", "A3"], [94, 158, 222], 70, "g"),
          ...makeGrid(["P1", "P2", "P3"], [94, 158, 222], 86, "w"),
          ...makeGrid(["Q1", "Q2", "Q3"], [94, 158, 222], 102, "g"),
          ...makeGrid(["S1", "S2", "S3"], [94, 158, 222], 118, "w"),
          ...makeGrid(["PF1", "PF2", "PF3"], [94, 158, 222], 134, "g"),
          { id: "P", x: 94, y: 153, bg: "w", scale: 0.96 },
          { id: "freq", alias: "f", x: 222, y: 153, bg: "w", scale: 0.96 },
          { id: "Q", x: 94, y: 169, bg: "g", scale: 0.96 },
          { id: "S", x: 94, y: 185, bg: "w", scale: 0.96 },
          { id: "PF", x: 94, y: 201, bg: "g", scale: 0.96, w_clear: 55 },
          { id: "An", x: 222, y: 201, bg: "g", scale: 0.96 }
        ].map(o => (["PF1", "PF2", "PF3"].includes(o.id) ? { ...o, w_clear: 55 } : o))
      },
      {
        id: "SD141",
        overlays: [
          { id: "V1", x: 63, y: 36, bg: "w", scale: 0.85, w_clear: 45 },
          { id: "Vdeg1", x: 121, y: 36, bg: "w", scale: 0.85, w_clear: 45 },
          { id: "V2", x: 63, y: 52, bg: "g", scale: 0.85, w_clear: 45 },
          { id: "Vdeg2", x: 121, y: 52, bg: "g", scale: 0.85, w_clear: 45 },
          { id: "V3", x: 63, y: 68, bg: "w", scale: 0.85, w_clear: 45 },
          { id: "Vdeg3", x: 121, y: 68, bg: "w", scale: 0.85, w_clear: 45 },
          { id: "A1", x: 63, y: 87, bg: "w", scale: 0.85, w_clear: 45 },
          { id: "Adeg1", x: 121, y: 87, bg: "w", scale: 0.85, w_clear: 45 },
          { id: "A2", x: 63, y: 103, bg: "g", scale: 0.85, w_clear: 45 },
          { id: "Adeg2", x: 121, y: 103, bg: "g", scale: 0.85, w_clear: 45 },
          { id: "A3", x: 63, y: 119, bg: "w", scale: 0.85, w_clear: 45 },
          { id: "Adeg3", x: 121, y: 119, bg: "w", scale: 0.85, w_clear: 45 },
          { id: "freq", alias: "f", x: 83, y: 154, bg: "w", scale: 0.85, w_clear: 45 },
          { id: "V_unb", alias: "V%", x: 83, y: 189, bg: "g", scale: 0.85, w_clear: 45 },
          { id: "A_unb", alias: "A%", x: 83, y: 205, bg: "w", scale: 0.85, w_clear: 45 }
        ]
      },
      { 
        id: "SD142", 
        overlays: [
          ...makeGrid(["V1", "V2", "V3"], [76, 136, 196], 47, "w"),
          ...makeGrid(["A1", "A2", "A3"], [76, 136, 196], 63, "g")
        ] 
      },
      {
        id: "SD143", overlays: [
          ...makeGrid(["V1", "V2", "V3"], [76, 136, 196], 47, "w"),
          ...makeGrid(["A1", "A2", "A3"], [76, 136, 196], 63, "g")
        ]
      },
      {
        id: "SD144", overlays: [
          ...makeGrid(["V1", "V2", "V3"], [76, 136, 196], 47, "w"),
          ...makeGrid(["THDV1", "THDV2", "THDV3"], [76, 136, 196], 63, "g")
        ]
      },
      {
        id: "SD145", overlays: [
          ...makeGrid(["A1", "A2", "A3"], [76, 136, 196], 47, "w"),
          ...makeGrid(["THDA1", "THDA2", "THDA3"], [76, 136, 196], 63, "g")
        ]
      }
    ];

    function makeGrid(ids, x_rights, y_bot, bg, scale = 0.96) {
      return ids.map((id, i) => ({ id, x: x_rights[i], y: y_bot, bg, scale }));
    }

    const DIGIT_TEMPLATES = {};
    const CHAR_MAP = { '.': 'dot', '-': 'minus', ':': 'colon' };

    async function loadDigitTemplates() {
      const symbols = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'dot', 'minus', 'colon'];
      const colors = ['w', 'g'];
      const promises = [];

      for (let s of symbols) {
        for (let c of colors) {
          const key = `${s}_${c}`;
          promises.push((async () => {
            const img = new Image();
            img.src = `templates/digits/${s}${c}.bmp`;
            await new Promise(r => { img.onload = r; img.onerror = r; });
            if (img.complete && img.naturalWidth > 0) {
              DIGIT_TEMPLATES[key] = img;
            }
          })());
        }
      }
      await Promise.all(promises);
    }

    function getDigitImg(char, color) {
      const s = CHAR_MAP[char] || char;
      let key = `${s}_${color}`;
      if (DIGIT_TEMPLATES[key]) return DIGIT_TEMPLATES[key];
      // Fallback
      key = `${s}_${color === 'w' ? 'g' : 'w'}`;
      return DIGIT_TEMPLATES[key] || null;
    }

    async function processImageClientSide(file, screenIdx, params) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = async (e) => {
          const img = new Image();
          img.onload = async () => {
            const canvas = document.createElement('canvas');
            canvas.width = img.width;
            canvas.height = img.height;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0);

            const templateSelect = document.getElementById('ei-template-select');
            const meterModel = templateSelect ? templateSelect.value : 'kew6315';
            
            let sc;
            if (meterModel === 'kew6315') {
                sc = SCREENS[screenIdx % 6] || SCREENS[0];
            } else {
                // Placeholder cho các mẫu đồng hồ khác (Hioki, Chauvin). Tạm thời fallback về SCREENS.
                sc = SCREENS[screenIdx % 6] || SCREENS[0];
            }
            
            for (const overlay of sc.overlays) {
              let val = params[overlay.id];
              if (val === undefined && overlay.alias) val = params[overlay.alias];
              if (val !== undefined && val !== null && val !== "") {
                await applyTextToCanvas(ctx, overlay, String(val));
              }
            }

            canvas.toBlob(blob => resolve(blob), 'image/bmp');
          };
          img.onerror = reject;
          img.src = e.target.result;
        };
        reader.onerror = reject;
        reader.readAsDataURL(file);
      });
    }

    async function applyTextToCanvas(ctx, config, text) {
      const x_right = config.x;
      const y_bot = config.y;
      const color = config.bg || 'w';
      const w_clear = config.w_clear || 50;
      const h_clear = 15;
      const scale = config.scale || 1.0;

      const x_left = Math.max(0, x_right - w_clear + 1);
      const y_top = Math.max(0, y_bot - h_clear + 1);

      // 1. Clear background (sample from bottom-left of clear area)
      const pixel = ctx.getImageData(x_left, y_bot, 1, 1).data;
      ctx.fillStyle = `rgb(${pixel[0]},${pixel[1]},${pixel[2]})`;
      ctx.fillRect(x_left, y_top, w_clear, h_clear);

      // 2. Draw text
      const chars = text.split('').reverse();
      let curr_x = x_right + 1;

      for (const char of chars) {
        let c = char === '/' ? '.' : char;
        const digitImg = getDigitImg(c, color);
        if (digitImg) {
          const dw = Math.floor(digitImg.naturalWidth * scale);
          const dh = Math.floor(digitImg.naturalHeight * scale);
          const spacing = dw >= 8 ? 1 : 2;

          curr_x -= dw;
          const pasteY = y_bot - dh + 1;
          ctx.drawImage(digitImg, curr_x, pasteY, dw, dh);
          curr_x -= spacing;
        } else {
          curr_x -= Math.floor(6 * scale);
        }
      }
    }



    // Development fallback (only runs if accessed as local file)
    if (window.location.protocol === 'file:') {
      document.getElementById('upload-container').style.display = 'none';
      document.getElementById('app-container').style.display = 'block';
      fetch('kew_analysis.json')
        .then(r => r.json())
        .then(data => renderAll(data));
    }
  

/* --------------------------------------------------- */