async function runSynopexReport() {
  const zipInput = document.getElementById('synopex-zip');
  const outputNameInput = document.getElementById('synopex-output-name');
  const statusEl = document.getElementById('synopex-status');
  const errorEl = document.getElementById('synopex-error');
  const btn = document.getElementById('btn-synopex-generate');
  const spinner = document.getElementById('synopex-spinner');
  const btnText = document.getElementById('synopex-btn-text');

  const zipFile = zipInput?.files?.[0];

  errorEl.style.display = 'none';
  errorEl.textContent = '';

  if (!zipFile) {
    errorEl.textContent = 'Vui lòng chọn file ZIP dữ liệu.';
    errorEl.style.display = 'block';
    return;
  }

  const formData = new FormData();
  formData.append('data_zip', zipFile);
  formData.append('output_name', (outputNameInput?.value || '').trim());

  statusEl.textContent = `Đang xử lý ${zipFile.name}...`;
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
  const zipInput = document.getElementById('synopex-zip');
  if (!zipInput) return;

  zipInput.addEventListener('change', () => {
    const statusEl = document.getElementById('synopex-status');
    if (zipInput.files?.[0] && statusEl) {
      statusEl.textContent = `Đã chọn ZIP: ${zipInput.files[0].name}`;
    }
  });
});
