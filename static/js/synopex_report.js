function updateSynopexFolderMeta() {
  const folderInput = document.getElementById('synopex-folder');
  const meta = document.getElementById('synopex-folder-meta');
  if (!folderInput || !meta) return;

  const files = Array.from(folderInput.files || []);
  if (files.length === 0) {
    meta.textContent = 'Chưa chọn thư mục dữ liệu.';
    return;
  }

  const roots = new Set(files.map(file => (file.webkitRelativePath || file.name).split('/')[0]).filter(Boolean));
  meta.textContent = `Đã chọn ${files.length} file từ ${roots.size} thư mục gốc.`;
}


async function runSynopexReport() {
  const templateInput = document.getElementById('synopex-template');
  const zipInput = document.getElementById('synopex-zip');
  const folderInput = document.getElementById('synopex-folder');
  const outputNameInput = document.getElementById('synopex-output-name');
  const tesseractInput = document.getElementById('synopex-tesseract');
  const statusEl = document.getElementById('synopex-status');
  const errorEl = document.getElementById('synopex-error');
  const btn = document.getElementById('btn-synopex-generate');
  const spinner = document.getElementById('synopex-spinner');
  const btnText = document.getElementById('synopex-btn-text');

  const templateFile = templateInput?.files?.[0];
  const zipFile = zipInput?.files?.[0];
  const folderFiles = Array.from(folderInput?.files || []);

  errorEl.style.display = 'none';
  errorEl.textContent = '';

  if (!templateFile) {
    errorEl.textContent = 'Vui lòng chọn file mẫu .docx.';
    errorEl.style.display = 'block';
    return;
  }

  if (!zipFile && folderFiles.length === 0) {
    errorEl.textContent = 'Vui lòng chọn dữ liệu nguồn bằng ZIP hoặc thư mục.';
    errorEl.style.display = 'block';
    return;
  }

  const formData = new FormData();
  formData.append('template', templateFile);
  formData.append('output_name', (outputNameInput?.value || '').trim());
  formData.append('tesseract_cmd', (tesseractInput?.value || '').trim());

  if (zipFile) {
    formData.append('data_zip', zipFile);
    statusEl.textContent = `Đang tạo báo cáo từ ZIP ${zipFile.name}...`;
  } else {
    for (const file of folderFiles) {
      formData.append('files', file, file.webkitRelativePath || file.name);
    }
    statusEl.textContent = `Đang tạo báo cáo từ ${folderFiles.length} file trong thư mục upload...`;
  }

  btn.disabled = true;
  spinner.style.display = 'inline-block';
  btnText.textContent = 'Đang xử lý...';

  try {
    const response = await fetch('/api/synopex/generate', {
      method: 'POST',
      body: formData,
    });

    if (!response.ok) {
      let message = `Lỗi HTTP ${response.status}`;
      try {
        const data = await response.json();
        if (data?.error) message = data.error;
      } catch (e) {
        // ignore
      }
      throw new Error(message);
    }

    const blob = await response.blob();
    const filenameHeader = response.headers.get('Content-Disposition') || '';
    const match = filenameHeader.match(/filename="?([^"]+)"?/i);
    const filename = match?.[1] || ((outputNameInput?.value || '').trim() || 'KEW_Synopex_Report.docx');
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    statusEl.textContent = `Đã tạo xong báo cáo: ${filename}`;
  } catch (error) {
    errorEl.textContent = error.message || 'Không thể tạo báo cáo.';
    errorEl.style.display = 'block';
    statusEl.textContent = '';
  } finally {
    btn.disabled = false;
    spinner.style.display = 'none';
    btnText.textContent = 'Tạo báo cáo Word';
  }
}


document.addEventListener('DOMContentLoaded', () => {
  const folderInput = document.getElementById('synopex-folder');
  const zipInput = document.getElementById('synopex-zip');

  if (folderInput) {
    folderInput.addEventListener('change', updateSynopexFolderMeta);
  }

  if (zipInput) {
    zipInput.addEventListener('change', () => {
      const statusEl = document.getElementById('synopex-status');
      if (zipInput.files?.[0] && statusEl) {
        statusEl.textContent = `Đã chọn ZIP: ${zipInput.files[0].name}`;
      }
    });
  }
});
