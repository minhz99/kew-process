let currentMode = 'string_mode';
let historyData = [];
let editingIndex = -1;
let workbook = null; // Store active workbook object
let sourceExcelFile = null; // Keep original uploaded file for server-side export
const pendingUpdateMap = new Map(); // key: `${sheet}::${address}` -> {sheet, address, value}
let currentFilename = "";

// Helper functions from Python logic
function to_number(val) {
    if (val === null || val === undefined) return null;
    if (typeof val === 'number') return val;
    let s = String(val).trim();
    if (!s) return null;

    // Clean Vietnamese/European number formatting (. as thousand separator, , as decimal or vice versa)
    s = s.replace(/\./g, '').replace(/,/g, '.');
    let n = parseFloat(s);
    return isNaN(n) ? val : n;
}


function find_row(ws, target_month, target_period) {
    // Read target cells (Column D/4 is Month, Column E/5 is Period)
    // ws indexing is A1, B1... or {c:3, r:4}
    let month_row = null;
    const range = XLSX.utils.decode_range(ws['!ref']);

    for (let r = 4; r <= range.e.r; r++) { // Row 5 is index 4
        let cellD = ws[XLSX.utils.encode_cell({ c: 3, r: r })];
        if (!cellD) continue;
        let valD = String(cellD.v).trim();
        if (valD == target_month || valD == target_month + ".0") {
            month_row = r;
            break;
        }
    }

    if (month_row === null) return (target_month - 1) * 4 + 4 + target_period;

    for (let r = month_row; r <= range.e.r; r++) {
        let cellE = ws[XLSX.utils.encode_cell({ c: 4, r: r })];
        let valE = cellE ? String(cellE.v).trim() : "";
        if (valE == target_period || valE == target_period + ".0") return r + 1; // Return 1-based row

        let cellD = ws[XLSX.utils.encode_cell({ c: 3, r: r })];
        let valD = cellD ? String(cellD.v).trim() : "";
        if (valD && valD !== "undefined" && r !== month_row) break;
        if (valE.toLowerCase() === "tổng") break;
    }
    return (target_month - 1) * 4 + 4 + target_period;
}

function buildCellUpdates(pairs, row) {
    const mapping = [
        { col: "F", val: pairs[0][1] }, { col: "G", val: pairs[0][0] },
        { col: "I", val: pairs[1][1] }, { col: "J", val: pairs[1][0] },
        { col: "L", val: pairs[2][1] }, { col: "M", val: pairs[2][0] }
    ];
    return mapping.map(m => ({
        address: m.col + row,
        value: typeof m.val === 'string' ? to_number(m.val) : m.val
    }));
}

function fill_excel(ws, pairs, row) {
    const updates = buildCellUpdates(pairs, row);
    updates.forEach(update => {
        if (!ws[update.address]) {
            ws[update.address] = { t: typeof update.value === 'number' ? 'n' : 's' };
        }
        ws[update.address].v = update.value;
    });
    return updates;
}

function registerPendingUpdates(sheetName, updates) {
    updates.forEach(update => {
        pendingUpdateMap.set(`${sheetName}::${update.address}`, {
            sheet: sheetName,
            address: update.address,
            value: update.value
        });
    });
}

function resetSessionEdits() {
    pendingUpdateMap.clear();
    historyData = [];
    editingIndex = -1;
    cancelEdit();
}

function parseFilenameFromContentDisposition(contentDisposition) {
    if (!contentDisposition) return null;
    const utf8Match = contentDisposition.match(/filename\*=UTF-8''([^;]+)/i);
    if (utf8Match && utf8Match[1]) {
        try {
            return decodeURIComponent(utf8Match[1]);
        } catch (_error) {
            return utf8Match[1];
        }
    }

    const asciiMatch = contentDisposition.match(/filename="?([^";]+)"?/i);
    return asciiMatch ? asciiMatch[1] : null;
}

// Xử lý chuyển tab
function switchExcelTab(modeId) {
    const root = document.getElementById('workspace-excel');
    if (!root) return;

    root.querySelectorAll('.tab-pane').forEach(el => el.classList.remove('active'));
    root.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));

    document.getElementById(modeId).classList.add('active');

    if (modeId === 'string_mode') {
        document.getElementById('tab_string').classList.add('active');
    } else {
        document.getElementById('tab_manual').classList.add('active');
    }
    currentMode = modeId;
}

// Hiện thông báo
function showMessage(msg, isError = false) {
    const msgDiv = document.getElementById('message');
    msgDiv.textContent = msg;
    msgDiv.className = isError ? 'error' : 'success';
}

// Lấy hoặc tạo session ID (để phân biệt các user/tab)
let sessionId = localStorage.getItem('dien_excel_session');
if (!sessionId) {
    sessionId = Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15);
    localStorage.setItem('dien_excel_session', sessionId);
}

// CLIENT-SIDE Upload File
async function uploadFile() {
    const fileInput = document.getElementById('excel_file');
    if (!fileInput.files.length) {
        alert('Vui lòng chọn ít nhất một tệp Excel (.xlsx).');
        return;
    }

    const file = fileInput.files[0];
    sourceExcelFile = file;
    currentFilename = file.name;
    resetSessionEdits();

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        workbook = XLSX.read(data, { type: 'array' });

        // Update UI
        document.getElementById('data_entry_section').style.opacity = '1';
        document.getElementById('data_entry_section').style.pointerEvents = 'auto';
        document.getElementById('active_file_display').style.display = 'flex';

        const selectEl = document.getElementById('active_filename');
        selectEl.innerHTML = `<option value="${currentFilename}">${currentFilename}</option>`;

        document.getElementById('btn_download').style.display = 'block';
        fileInput.value = '';
        showMessage(`Đã tải file ${currentFilename} thành công!`);
    };
    reader.readAsArrayBuffer(file);
}

// SERVER-SIDE Download File (preserve workbook formatting)
async function downloadFile() {
    if (!sourceExcelFile) {
        showMessage("Vui lòng tải file Excel lên trước.", true);
        return;
    }

    const updates = Array.from(pendingUpdateMap.values());
    if (updates.length === 0) {
        showMessage("Chưa có dữ liệu nào để xuất file.", true);
        return;
    }

    const btnDownload = document.getElementById('btn_download');
    const originalLabel = btnDownload ? btnDownload.textContent : null;
    if (btnDownload) {
        btnDownload.disabled = true;
        btnDownload.textContent = 'Đang tạo file...';
    }

    try {
        const formData = new FormData();
        formData.append('file', sourceExcelFile, currentFilename || sourceExcelFile.name);
        formData.append('updates', JSON.stringify(updates));
        formData.append('filename', currentFilename || sourceExcelFile.name || `KetQua_Excel_${new Date().getTime()}.xlsx`);

        const response = await fetch('/api/excel/apply-updates', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            let errorMessage = 'Không thể tạo file Excel kết quả.';
            try {
                const payload = await response.json();
                if (payload && payload.error) errorMessage = payload.error;
            } catch (_error) {
                // ignore parse errors and use fallback error message
            }
            throw new Error(errorMessage);
        }

        const blob = await response.blob();
        const headerName = parseFilenameFromContentDisposition(response.headers.get('Content-Disposition') || '');
        const outputName = headerName || currentFilename || `KetQua_Excel_${new Date().getTime()}.xlsx`;

        const downloadUrl = URL.createObjectURL(blob);
        const anchor = document.createElement('a');
        anchor.href = downloadUrl;
        anchor.download = outputName;
        document.body.appendChild(anchor);
        anchor.click();
        anchor.remove();
        URL.revokeObjectURL(downloadUrl);

        showMessage("✓ Đã tạo file Excel thành công, định dạng ô được giữ nguyên.");
    } catch (error) {
        showMessage(error.message || "Xuất file thất bại.", true);
    } finally {
        if (btnDownload) {
            btnDownload.disabled = false;
            btnDownload.textContent = originalLabel;
        }
    }
}

// Kéo thả file vào màn hình để upload
const dropZone = document.body;
const dragOverlay = document.createElement('div');
dragOverlay.id = 'drag_overlay';
dragOverlay.innerHTML = '<h2>Kéo thả file(s) Excel vào đây để tải lên</h2>';
dragOverlay.style.cssText = `
            display: none;
            position: fixed;
            top: 0; left: 0; right: 0; bottom: 0;
            background: rgba(10, 18, 34, 0.94);
            border: 2px dashed rgba(34, 197, 94, 0.7);
            color: #dcfce7;
            z-index: 9999;
            align-items: center;
            justify-content: center;
            text-align: center;
        `;
document.body.appendChild(dragOverlay);

function isExcelWorkspaceActive() {
    const excelWorkspace = document.getElementById('workspace-excel');
    return excelWorkspace && excelWorkspace.style.display !== 'none';
}

function hasFilesInDragEvent(e) {
    return e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files');
}

let dragCounter = 0;
dropZone.addEventListener('dragenter', (e) => {
    e.preventDefault();
    if (!isExcelWorkspaceActive() || !hasFilesInDragEvent(e)) return;
    dragCounter++;
    dragOverlay.style.display = 'flex';
});
dropZone.addEventListener('dragleave', (e) => {
    e.preventDefault();
    if (!isExcelWorkspaceActive()) return;
    dragCounter--;
    if (dragCounter === 0) dragOverlay.style.display = 'none';
});
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    if (!isExcelWorkspaceActive()) return;
});
dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    if (!isExcelWorkspaceActive()) return;
    dragCounter = 0;
    dragOverlay.style.display = 'none';
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        document.getElementById('excel_file').files = e.dataTransfer.files;
        uploadFile();
    }
});

// Xử lý Paste cho tất cả các field Manual (Trợ giúp nhanh cho Manual Mode)
document.getElementById('bt_price').addEventListener('paste', function (event) {
    const pasteData = (event.clipboardData || window.clipboardData).getData('text');
    const lines = pasteData.trim().split(/\r?\n/).map(line => line.trim()).filter(line => line.length > 0);

    if (lines.length >= 3) {
        event.preventDefault();
        const parseLine = (line) => line.split(/\s+/);
        const bt = parseLine(lines[0]), cd = parseLine(lines[1]), td = parseLine(lines[2]);

        if (bt.length >= 2) { document.getElementById('bt_price').value = bt[0]; document.getElementById('bt_usage').value = bt[1]; }
        if (cd.length >= 2) { document.getElementById('cd_price').value = cd[0]; document.getElementById('cd_usage').value = cd[1]; }
        if (td.length >= 2) { document.getElementById('td_price').value = td[0]; document.getElementById('td_usage').value = td[1]; }
    }
});

// CLIENT-SIDE Submit Dữ Liệu
async function submitData() {
    if (!workbook) {
        alert("Vui lòng tải file Excel lên trước.");
        return;
    }

    const sheetName = document.getElementById('sheet_name').value.trim();
    const month = parseInt(document.getElementById('month').value);
    const period = parseInt(document.getElementById('period').value);

    if (!workbook.Sheets[sheetName]) {
        showMessage(`Không tìm thấy sheet "${sheetName}"`, true);
        return;
    }

    const ws = workbook.Sheets[sheetName];
    let parsed_groups = [];

    function format_val(v) {
        v = v.trim();
        if (/^\d{4,}$/.test(v)) return v.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
        return v;
    }

    if (currentMode === 'string_mode') {
        const rawText = document.getElementById('raw_data').value;
        const lines = rawText.split(/\n/).filter(l => l.trim().length > 0);

        // Simple Parser (3 lines logic)
        for (let i = 0; i < lines.length; i += 3) {
            if (i + 2 >= lines.length) break;
            const items = [lines[i], lines[i + 1], lines[i + 2]].map(l => {
                const nums = l.match(/\d+[\d\.,]*/g) || [];
                let p = "0", q = "0";
                if (nums.length >= 3) {
                    p = nums[nums.length - 3];
                    q = nums[nums.length - 2];
                } else if (nums.length === 2) {
                    p = nums[0];
                    q = nums[1];
                } else if (nums.length === 1) {
                    p = nums[0];
                }
                return [format_val(p), format_val(q)];
            });
            parsed_groups.push(items);
        }
    } else {
        parsed_groups = [[
            [format_val(document.getElementById('bt_price').value), format_val(document.getElementById('bt_usage').value)],
            [format_val(document.getElementById('cd_price').value), format_val(document.getElementById('cd_usage').value)],
            [format_val(document.getElementById('td_price').value), format_val(document.getElementById('td_usage').value)]
        ]];
    }

    if (parsed_groups.length === 0) {
        showMessage("Dữ liệu không hợp lệ.", true);
        return;
    }

    const startRow = find_row(ws, month, period);
    let inserted_results = [];

    parsed_groups.forEach((group, i) => {
        const targetRow = startRow + i;
        const updates = fill_excel(ws, group, targetRow);
        registerPendingUpdates(sheetName, updates);
        inserted_results.push({ row: targetRow, parsed_data: group });
    });

    // Update UI History
    const timeStr = new Date().toLocaleTimeString('vi-VN');
    inserted_results.reverse().forEach(g => {
        historyData.unshift({
            time: timeStr,
            filename: currentFilename,
            sheet: sheetName,
            month: month,
            period: period,
            parsed_data: g.parsed_data,
            row: g.row
        });
    });

    if (document.getElementById('auto_increment').checked) {
        let m = month, p = period;
        if (p === 3) { p = 1; m = m < 12 ? m + 1 : 1; } else { p++; }
        document.getElementById('month').value = m;
        document.getElementById('period').value = p;
    }

    renderHistoryTable();
    showMessage("✓ Đã ghi dữ liệu thành công! Có thể tải file xuống và giữ nguyên định dạng ô.");
}

// Bảng History
function renderHistoryTable() {
    const tbody = document.getElementById('history_body');
    tbody.innerHTML = '';

    if (historyData.length === 0) {
        tbody.innerHTML = `<tr><td colspan="8" class="empty-state">Chưa có kết quả nào được ghi trong phiên này.</td></tr>`;
        return;
    }

    historyData.forEach((item, index) => {
        const tr = document.createElement('tr');
        if (index === editingIndex) {
            tr.className = 'row-edited';
        }

        const formatCell = (pair) => `<b>${pair[0]}</b> - ${pair[1]}`;
        const fileDisp = item.filename.length > 20 ? item.filename.substring(0, 17) + '...' : item.filename;

        tr.innerHTML = `
                    <td>${item.time}</td>
                    <td title="${item.filename}"><b>${fileDisp}</b></td>
                    <td><b>${item.month}</b></td>
                    <td><b>${item.period}</b></td>
                    <td>${formatCell(item.parsed_data[0])}</td>
                    <td>${formatCell(item.parsed_data[1])}</td>
                    <td>${formatCell(item.parsed_data[2])}</td>
                    <td style="text-align: right;">
                        <button class="btn-small btn-warning" onclick="editRow(${index})">🛠 Sửa</button>
                    </td>
                `;
        tbody.appendChild(tr);
    });
}

// Nhấn nút Edit trên lịch sử
function editRow(index) {
    editingIndex = index;
    const item = historyData[index];

    // Chọn đúng file trong dropdown
    const selectEl = document.getElementById('active_filename');
    for (let i = 0; i < selectEl.options.length; i++) {
        if (selectEl.options[i].value === item.filename) {
            selectEl.value = item.filename;
            break;
        }
    }

    // Điền lại form
    document.getElementById('sheet_name').value = item.sheet;
    document.getElementById('month').value = item.month;
    document.getElementById('period').value = item.period;

    const [bt, cd, td] = item.parsed_data;

    // Điền dữ liệu Manual (LUÔN update Manual Mode để có sẵn)
    document.getElementById('bt_price').value = bt[0]; document.getElementById('bt_usage').value = bt[1];
    document.getElementById('cd_price').value = cd[0]; document.getElementById('cd_usage').value = cd[1];
    document.getElementById('td_price').value = td[0]; document.getElementById('td_usage').value = td[1];

    // Lắp ráp lại chuỗi Textarea
    document.getElementById('raw_data').value = `${bt[0]} ${bt[1]}\
${cd[0]} ${cd[1]}\
${td[0]} ${td[1]}`;

    // Chuyển UI nút Submit
    const btnSubmit = document.getElementById('btn_submit');
    btnSubmit.textContent = `LƯU VÀ GHI ĐÈ EXCEL (KỲ ${item.period} THÁNG ${item.month} NĂM ${item.sheet})`;
    btnSubmit.className = "btn-block btn-warning";

    document.getElementById('btn_cancel_edit').style.display = 'block';

    // Trượt lên form
    document.getElementById('data_entry_section').scrollIntoView({ behavior: 'smooth' });

    renderHistoryTable(); // Highlight row
}

function cancelEdit() {
    editingIndex = -1;

    const btnSubmit = document.getElementById('btn_submit');
    btnSubmit.textContent = "GHI DỮ LIỆU VÀO EXCEL";
    btnSubmit.className = "btn-block";

    document.getElementById('btn_cancel_edit').style.display = 'none';
    renderHistoryTable(); // Remove highlight
}

// Bắt sự kiện Enter global
document.addEventListener('keydown', function (event) {
    if (event.ctrlKey && event.key === 'Enter') {
        submitData();
    }
});
document.getElementById('raw_data').addEventListener('keydown', function (event) {
    if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        submitData();
    }
});
