// ===== Global State =====
const state = {
    allPersons: [],      // All data
    personsByBus: {},    // Data grouped by bus
    currentBus: 'A',     // Current selected bus
    selectedIndex: null,
    splitCount: 4,
    isApplied: false,    // Whether format is applied
    // Shared config (same for all buses)
    config: {
        companyName: '台灣通力電梯',
        eventName: '親清日月潭山系',
        travelInfo: 'FOUND WONDER TRAVEL 農新家旅遊',
        busLabel: '★ 車 / 船號：',
        tableLabel: '★ 桌　　號：',
        roomLabel: '★ 房　　號：'
    },
    // Font sizes for 6 groups (shared across all buses)
    fontSizes: {
        company: 3.5,      // 公司名稱 (mm)
        event: 5,          // 活動名稱 (mm)
        name: 12,          // 姓名 (mm) - range 12-40
        travel: 2.8,       // 旅行社資訊 (mm)
        labels: 3.5,       // 標籤區 (mm)
        footer: 3          // 底部資訊 (mm)
    },
    // Per-bus settings (colors and footer)
    busSetting: {
        'A': { bgColor: '#d4e6d4', textColor: '#2a4a2a', borderColor: '#5a7a5a', footerText: 'A 車 領隊 黃啟翔 0935-670-825' },
        'B': { bgColor: '#d4e0e6', textColor: '#2a3a4a', borderColor: '#5a6a7a', footerText: 'B 車 領隊 待確認' },
        'C': { bgColor: '#e6e4d4', textColor: '#4a4a2a', borderColor: '#7a7a5a', footerText: 'C 車 領隊 待確認' },
        'D': { bgColor: '#e6d4d4', textColor: '#4a2a2a', borderColor: '#7a5a5a', footerText: 'D 車 領隊 待確認' },
        'E': { bgColor: '#e4d4e6', textColor: '#3a2a4a', borderColor: '#6a5a7a', footerText: 'E 車 領隊 待確認' }
    }
};

// ===== DOM Elements =====
const elements = {
    appContainer: document.getElementById('appContainer'),
    uploadZone: document.getElementById('uploadZone'),
    fileInput: document.getElementById('fileInput'),
    fileInfo: document.getElementById('fileInfo'),
    busTabs: document.getElementById('busTabs'),
    previewGrid: document.getElementById('previewGrid'),
    previewCount: document.getElementById('previewCount'),
    previewModeInfo: document.getElementById('previewModeInfo'),
    editModal: document.getElementById('editModal'),
    toast: document.getElementById('toast'),
    toastMessage: document.getElementById('toastMessage'),
    applyBtn: document.getElementById('applyBtn'),
    exportPdf: document.getElementById('exportPdf'),
    exportSelectedPdf: document.getElementById('exportSelectedPdf'),
    exportWord: document.getElementById('exportWord'),
    exportSelectedWord: document.getElementById('exportSelectedWord'),
    printBtn: document.getElementById('printBtn'),
    busCheckboxes: document.getElementById('busCheckboxes'),
    selectAllBuses: document.getElementById('selectAllBuses'),
    deselectAllBuses: document.getElementById('deselectAllBuses')
};

// ===== Debounce utility =====
function debounce(func, wait) {
    let timeout;
    return function(...args) {
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(this, args), wait);
    };
}

// ===== Toast Function =====
function showToast(message, type = 'success') {
    elements.toastMessage.textContent = message;
    elements.toast.className = `toast ${type}`;
    elements.toast.classList.add('show');

    setTimeout(() => {
        elements.toast.classList.remove('show');
    }, 3000);
}

// ===== File Import Functions =====
function handleFileUpload(file) {
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Read from 總表 sheet
            const sheetName = workbook.SheetNames.includes('總表') ? '總表' : workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(sheet);

            state.allPersons = jsonData.map((row, index) => {
                // Get table value - could be number or "素"
                let tableValue = row['桌次&活動'] || row['桌次'] || row['桌號'] || '';
                let isVeg = false;

                if (tableValue === '素' || String(tableValue).includes('素')) {
                    isVeg = true;
                    tableValue = String(tableValue).replace('素', '').trim() || '素食桌';
                }

                // Get bus value - support multiple column names
                let busValue = row['車次'] || row['車 / 船號'] || row['車次&遊湖'] || row['車號'] || 'A';
                busValue = String(busValue).charAt(0).toUpperCase();

                return {
                    id: index,
                    name: row['姓名'] || '',
                    bus: busValue,
                    table: String(tableValue),
                    room: String(row['房號'] || ''),
                    note: row['備註'] || '',
                    isVeg: isVeg,
                    override: {
                        fontSize: null,
                        bgColor: null
                    }
                };
            }).filter(p => p.name); // Filter out empty rows

            // Group by bus
            state.personsByBus = { A: [], B: [], C: [], D: [], E: [] };
            state.allPersons.forEach(person => {
                const bus = person.bus.charAt(0).toUpperCase();
                if (state.personsByBus[bus]) {
                    state.personsByBus[bus].push(person);
                }
            });

            // Update UI
            updateBusTabs();
            state.isApplied = false;
            renderPreview();

            elements.uploadZone.classList.add('has-file');
            elements.fileInfo.style.display = 'block';
            elements.fileInfo.textContent = `已匯入：${file.name} (${state.allPersons.length} 筆資料)`;
            elements.busTabs.style.display = 'flex';
            elements.applyBtn.disabled = false;

            showToast(`成功匯入 ${state.allPersons.length} 筆資料`, 'success');
        } catch (error) {
            console.error('File parse error:', error);
            showToast('檔案解析失敗，請確認格式正確', 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

function updateBusTabs() {
    document.querySelectorAll('.bus-tab').forEach(tab => {
        const bus = tab.dataset.bus;
        const count = state.personsByBus[bus]?.length || 0;
        tab.querySelector('.count').textContent = count;

        // Hide tabs with no data
        tab.style.display = count > 0 ? 'flex' : 'none';
    });

    // Update checkbox counts in export panel
    document.querySelectorAll('.check-count').forEach(span => {
        const bus = span.dataset.bus;
        const count = state.personsByBus[bus]?.length || 0;
        span.textContent = count;
    });

    // Select first bus with data
    const firstBusWithData = ['A', 'B', 'C', 'D', 'E'].find(bus =>
        state.personsByBus[bus]?.length > 0
    );
    if (firstBusWithData) {
        selectBus(firstBusWithData);
    }
}

function selectBus(bus) {
    state.currentBus = bus;
    state.isApplied = false;

    document.querySelectorAll('.bus-tab').forEach(tab => {
        tab.classList.toggle('active', tab.dataset.bus === bus);
    });

    // Update current bus label
    const busLabel = document.getElementById('currentBusLabel');
    if (busLabel) busLabel.textContent = bus;

    // Load bus-specific settings
    const busSetting = state.busSetting[bus];
    if (busSetting) {
        document.getElementById('footerText').value = busSetting.footerText;
        document.getElementById('bgColor').value = busSetting.bgColor;
        document.getElementById('textColor').value = busSetting.textColor;
        document.getElementById('borderColor').value = busSetting.borderColor;
    }

    renderPreview();
}

function setupUploadZone() {
    elements.uploadZone.addEventListener('click', () => {
        elements.fileInput.click();
    });

    elements.fileInput.addEventListener('change', (e) => {
        handleFileUpload(e.target.files[0]);
    });

    elements.uploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        elements.uploadZone.classList.add('dragover');
    });

    elements.uploadZone.addEventListener('dragleave', () => {
        elements.uploadZone.classList.remove('dragover');
    });

    elements.uploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        elements.uploadZone.classList.remove('dragover');
        handleFileUpload(e.dataTransfer.files[0]);
    });
}

function clearData() {
    if (state.allPersons.length === 0) {
        showToast('目前沒有資料', 'error');
        return;
    }
    if (confirm('確定要清除所有資料嗎？')) {
        state.allPersons = [];
        state.personsByBus = { A: [], B: [], C: [], D: [], E: [] };
        state.isApplied = false;
        renderPreview();
        elements.uploadZone.classList.remove('has-file');
        elements.fileInfo.style.display = 'none';
        elements.busTabs.style.display = 'none';
        elements.fileInput.value = '';
        elements.applyBtn.disabled = true;
        elements.exportPdf.disabled = true;
        elements.exportSelectedPdf.disabled = true;
        elements.exportWord.disabled = true;
        elements.exportSelectedWord.disabled = true;
        elements.printBtn.disabled = true;
        showToast('資料已清除', 'success');
    }
}

// ===== Preview Functions =====
function renderPreview() {
    const persons = state.personsByBus[state.currentBus] || [];

    if (persons.length === 0) {
        elements.previewGrid.innerHTML = `
            <div class="empty-state">
                <div class="empty-icon">📋</div>
                <div class="empty-text">尚未匯入資料</div>
                <div class="empty-hint">請從左側面板上傳 Excel 或 CSV 檔案</div>
            </div>
        `;
        elements.previewCount.textContent = '共 0 張名牌';
        elements.previewModeInfo.style.display = 'none';
        return;
    }

    const split = state.splitCount;
    const totalPages = Math.ceil(persons.length / split);

    // If not applied, show first page with actual data (up to split count)
    if (!state.isApplied) {
        elements.previewModeInfo.style.display = 'block';
        const previewCount = Math.min(split, persons.length);
        elements.previewCount.textContent = `${state.currentBus}車：共 ${persons.length} 張，預覽第 1 頁（${previewCount} 張）`;

        let html = `<div class="a4-page split-${split}">`;

        // Show first page of nametags
        for (let i = 0; i < split; i++) {
            if (i < persons.length) {
                const person = persons[i];
                html += renderNametagHTML(person, i);
            } else {
                // Empty slot
                html += `<div class="nametag" style="background: #fff; border-color: #ddd;"></div>`;
            }
        }
        html += '</div>';

        elements.previewGrid.innerHTML = html;
    } else {
        // Show all pages
        elements.previewModeInfo.style.display = 'none';

        // Use DocumentFragment for better performance
        const fragment = document.createDocumentFragment();

        for (let pageIndex = 0; pageIndex < totalPages; pageIndex++) {
            const pageDiv = document.createElement('div');
            pageDiv.className = `a4-page split-${split}`;

            for (let i = 0; i < split; i++) {
                const personIndex = pageIndex * split + i;
                if (personIndex < persons.length) {
                    const person = persons[personIndex];
                    pageDiv.innerHTML += renderNametagHTML(person, personIndex);
                } else {
                    pageDiv.innerHTML += `<div class="nametag" style="background: #fff; border-color: #ddd;"></div>`;
                }
            }
            fragment.appendChild(pageDiv);
        }

        elements.previewGrid.innerHTML = '';
        elements.previewGrid.appendChild(fragment);
        elements.previewCount.textContent = `${state.currentBus}車：${persons.length} 張名牌，${totalPages} 頁 A4`;
    }

    // Add click listeners using event delegation
    elements.previewGrid.onclick = function(e) {
        const card = e.target.closest('.nametag[data-index]');
        if (card) {
            const index = parseInt(card.dataset.index);
            openEditModal(index);
        }
    };
}

// Debounced version for input events
const debouncedRenderPreview = debounce(renderPreview, 100);

function renderNametagHTML(person, index) {
    const isVeg = person.isVeg || (person.note && person.note.includes('素'));

    // Get bus-specific colors
    const busSetting = state.busSetting[person.bus] || state.busSetting['A'];
    const bgColor = person.override.bgColor || busSetting.bgColor;
    const textColor = busSetting.textColor;
    const borderColor = busSetting.borderColor;
    const footerText = busSetting.footerText;

    // Get font sizes (use person override for name if set)
    const nameFontSize = person.override.fontSize || state.fontSizes.name;

    return `
        <div class="nametag" data-index="${index}" style="background: ${bgColor}; border-color: ${borderColor};">
            <div class="nametag-header">
                <div class="nametag-company" style="font-size: ${state.fontSizes.company}mm; color: ${borderColor};">${state.config.companyName}</div>
                <div class="nametag-main-title" style="font-size: ${state.fontSizes.event}mm; color: ${textColor};">${state.config.eventName}</div>
                <div class="nametag-sub-info" style="font-size: ${state.fontSizes.travel}mm; color: ${borderColor};">${state.config.travelInfo}</div>
            </div>
            <div class="nametag-name" style="font-size: ${nameFontSize}mm; color: ${textColor}; border-color: ${borderColor};">
                ${person.name}
            </div>
            <div class="nametag-details" style="font-size: ${state.fontSizes.labels}mm;">
                <div class="nametag-detail-row">
                    <span class="detail-label" style="color: ${borderColor};">${state.config.busLabel}</span>
                    <span class="detail-value" style="color: ${textColor};">${person.bus} 車</span>
                </div>
                <div class="nametag-detail-row">
                    <span class="detail-label" style="color: ${borderColor};">${state.config.tableLabel}</span>
                    <span class="detail-value" style="color: ${textColor};">${person.table}</span>
                    ${isVeg ? `<span class="veg-tag">素食桌</span>` : ''}
                </div>
                <div class="nametag-detail-row">
                    <span class="detail-label" style="color: ${borderColor};">${state.config.roomLabel}</span>
                    <span class="detail-value" style="color: ${textColor};">${person.room}</span>
                </div>
            </div>
            <div class="nametag-footer" style="font-size: ${state.fontSizes.footer}mm; color: ${borderColor}; border-color: ${borderColor};">
                ${footerText}
            </div>
        </div>
    `;
}

// ===== Apply Format =====
function applyFormat() {
    state.isApplied = true;
    elements.exportPdf.disabled = false;
    elements.exportSelectedPdf.disabled = false;
    elements.exportWord.disabled = false;
    elements.exportSelectedWord.disabled = false;
    elements.printBtn.disabled = false;
    renderPreview();
    showToast('格式已套用，顯示全部名牌', 'success');
}

// ===== Config Update Functions =====
function setupConfigListeners() {
    // Template settings - use debounced render for text inputs (shared across all buses)
    ['companyName', 'eventName', 'busLabel', 'tableLabel', 'roomLabel'].forEach(id => {
        document.getElementById(id).addEventListener('input', (e) => {
            state.config[id] = e.target.value;
            debouncedRenderPreview();
        });
    });

    // Footer text - save per bus
    document.getElementById('footerText').addEventListener('input', (e) => {
        state.busSetting[state.currentBus].footerText = e.target.value;
        debouncedRenderPreview();
    });

    // Color settings - save per bus
    ['bgColor', 'textColor', 'borderColor'].forEach(id => {
        document.getElementById(id).addEventListener('input', (e) => {
            state.busSetting[state.currentBus][id] = e.target.value;
            debouncedRenderPreview();
        });
    });

    // Font size sliders - 6 groups (shared across all buses)
    const fontSizeMap = {
        'fontSizeCompany': 'company',
        'fontSizeEvent': 'event',
        'fontSizeName': 'name',
        'fontSizeTravel': 'travel',
        'fontSizeLabels': 'labels',
        'fontSizeFooter': 'footer'
    };

    Object.keys(fontSizeMap).forEach(sliderId => {
        const slider = document.getElementById(sliderId);
        if (slider) {
            slider.addEventListener('input', (e) => {
                const size = parseFloat(e.target.value);
                state.fontSizes[fontSizeMap[sliderId]] = size;
                // Update display value
                const valueSpan = document.querySelector(`.font-size-value[data-target="${sliderId}"]`);
                if (valueSpan) valueSpan.textContent = `${size}mm`;
                debouncedRenderPreview();
            });
        }
    });

    // Size presets (A4 split)
    document.querySelectorAll('.size-preset').forEach(preset => {
        preset.addEventListener('click', () => {
            document.querySelectorAll('.size-preset').forEach(p => p.classList.remove('active'));
            preset.classList.add('active');
            state.splitCount = parseInt(preset.dataset.split);
            renderPreview();
        });
    });

    // Bus tabs
    document.querySelectorAll('.bus-tab').forEach(tab => {
        tab.addEventListener('click', () => {
            selectBus(tab.dataset.bus);
        });
    });

    // Apply button
    elements.applyBtn.addEventListener('click', applyFormat);
}

// ===== Edit Modal Functions =====
function openEditModal(index) {
    const persons = state.personsByBus[state.currentBus] || [];
    if (index >= persons.length) return;

    state.selectedIndex = index;
    const person = persons[index];
    const busSetting = state.busSetting[person.bus] || state.busSetting['A'];

    document.getElementById('editName').value = person.name;
    document.getElementById('editBus').value = person.bus;
    document.getElementById('editTable').value = person.table;
    document.getElementById('editRoom').value = person.room;
    document.getElementById('editNote').value = person.note || '';
    document.getElementById('editFontSize').value = person.override.fontSize || state.fontSizes.name;
    document.getElementById('editBgColor').value = person.override.bgColor || busSetting.bgColor;

    elements.editModal.classList.add('active');
}

function closeEditModal() {
    elements.editModal.classList.remove('active');
    state.selectedIndex = null;
}

function saveEdit() {
    if (state.selectedIndex === null) return;

    const persons = state.personsByBus[state.currentBus];
    const person = persons[state.selectedIndex];

    person.name = document.getElementById('editName').value;
    person.bus = document.getElementById('editBus').value;
    person.table = document.getElementById('editTable').value;
    person.room = document.getElementById('editRoom').value;
    person.note = document.getElementById('editNote').value;
    person.isVeg = person.note && person.note.includes('素');
    person.override.fontSize = parseInt(document.getElementById('editFontSize').value);
    person.override.bgColor = document.getElementById('editBgColor').value;

    renderPreview();
    closeEditModal();
    showToast('名牌已更新', 'success');
}

function setupModalListeners() {
    document.getElementById('closeModal').addEventListener('click', closeEditModal);
    document.getElementById('cancelEdit').addEventListener('click', closeEditModal);
    document.getElementById('saveEdit').addEventListener('click', saveEdit);

    elements.editModal.addEventListener('click', (e) => {
        if (e.target === elements.editModal) {
            closeEditModal();
        }
    });
}

// ===== Panel Toggle =====
function setupPanelToggles() {
    document.querySelectorAll('.panel-header').forEach(header => {
        header.addEventListener('click', () => {
            header.closest('.panel').classList.toggle('collapsed');
        });
    });
}

// ===== Export Functions =====
async function exportPdf() {
    if (!state.isApplied) {
        showToast('請先點擊「套用格式並預覽全部」', 'error');
        return;
    }

    showToast('正在生成 PDF...', 'success');

    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({
        orientation: 'portrait',
        unit: 'mm',
        format: 'a4'
    });

    const pages = document.querySelectorAll('.a4-page');

    for (let i = 0; i < pages.length; i++) {
        if (i > 0) {
            pdf.addPage();
        }

        const canvas = await html2canvas(pages[i], {
            scale: 2,
            useCORS: true,
            backgroundColor: '#ffffff'
        });

        const imgData = canvas.toDataURL('image/jpeg', 0.95);
        pdf.addImage(imgData, 'JPEG', 0, 0, 210, 297);
    }

    pdf.save(`名牌輸出_${state.currentBus}車.pdf`);
    showToast('PDF 已下載', 'success');
}

// Get selected buses from checkboxes
function getSelectedBuses() {
    const checkboxes = elements.busCheckboxes.querySelectorAll('input[type="checkbox"]:checked');
    return Array.from(checkboxes).map(cb => cb.value);
}

async function exportSelectedPdf() {
    if (!state.isApplied) {
        showToast('請先點擊「套用格式並預覽全部」', 'error');
        return;
    }

    const selectedBuses = getSelectedBuses();

    // Filter buses that have data
    const busesWithData = selectedBuses.filter(bus =>
        state.personsByBus[bus] && state.personsByBus[bus].length > 0
    );

    if (busesWithData.length === 0) {
        showToast('所選車次中沒有資料', 'error');
        return;
    }

    showToast(`正在生成 ${busesWithData.length} 個車次 PDF...`, 'success');

    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({
        orientation: 'portrait',
        unit: 'mm',
        format: 'a4'
    });

    let isFirstPage = true;
    const originalBus = state.currentBus;

    for (const bus of busesWithData) {
        const persons = state.personsByBus[bus];
        if (!persons || persons.length === 0) continue;

        // Temporarily switch to this bus and render
        state.currentBus = bus;
        state.config.footerText = state.footerByBus[bus] || `${bus} 車 領隊 待確認`;
        renderPreview();

        // Wait for render
        await new Promise(resolve => setTimeout(resolve, 100));

        const pages = document.querySelectorAll('.a4-page');

        for (let i = 0; i < pages.length; i++) {
            if (!isFirstPage) {
                pdf.addPage();
            }
            isFirstPage = false;

            const canvas = await html2canvas(pages[i], {
                scale: 2,
                useCORS: true,
                backgroundColor: '#ffffff'
            });

            const imgData = canvas.toDataURL('image/jpeg', 0.95);
            pdf.addImage(imgData, 'JPEG', 0, 0, 210, 297);
        }
    }

    // Restore original bus
    state.currentBus = originalBus;
    selectBus(originalBus);

    const busLabels = busesWithData.join('');
    pdf.save(`名牌輸出_${busLabels}車.pdf`);
    showToast(`${busesWithData.length} 個車次 PDF 已下載`, 'success');
}

function printNametags() {
    if (!state.isApplied) {
        showToast('請先點擊「套用格式並預覽全部」', 'error');
        return;
    }
    window.print();
}

// ===== Word Export Functions =====
async function exportWord() {
    if (!state.isApplied) {
        showToast('請先點擊「套用格式並預覽全部」', 'error');
        return;
    }

    showToast('正在生成 Word 文件...', 'success');

    try {
        const doc = await generateWordDocument([state.currentBus]);
        const blob = await docx.Packer.toBlob(doc);
        saveAs(blob, `名牌輸出_${state.currentBus}車.docx`);
        showToast('Word 文件已下載', 'success');
    } catch (error) {
        console.error('Word export error:', error);
        showToast('Word 輸出失敗', 'error');
    }
}

async function exportSelectedWord() {
    if (!state.isApplied) {
        showToast('請先點擊「套用格式並預覽全部」', 'error');
        return;
    }

    const selectedBuses = getSelectedBuses();
    const busesWithData = selectedBuses.filter(bus =>
        state.personsByBus[bus] && state.personsByBus[bus].length > 0
    );

    if (busesWithData.length === 0) {
        showToast('所選車次中沒有資料', 'error');
        return;
    }

    showToast(`正在生成 ${busesWithData.length} 個車次 Word 文件...`, 'success');

    try {
        const doc = await generateWordDocument(busesWithData);
        const blob = await docx.Packer.toBlob(doc);
        const busLabels = busesWithData.join('');
        saveAs(blob, `名牌輸出_${busLabels}車.docx`);
        showToast(`${busesWithData.length} 個車次 Word 文件已下載`, 'success');
    } catch (error) {
        console.error('Word export error:', error);
        showToast('Word 輸出失敗', 'error');
    }
}

async function generateWordDocument(buses) {
    const { Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle, PageOrientation } = docx;

    // Helper function to convert mm to half-points (Word uses half-points)
    const mmToHalfPoints = (mm) => Math.round(mm * 2.835 * 2);

    // Calculate grid based on split count
    let cols, rows;
    switch (state.splitCount) {
        case 1: cols = 1; rows = 1; break;
        case 2: cols = 1; rows = 2; break;
        case 4: cols = 2; rows = 2; break;
        case 6: cols = 2; rows = 3; break;
        case 9: cols = 3; rows = 3; break;
        default: cols = 2; rows = 2;
    }

    const cellsPerPage = cols * rows;
    const sections = [];

    for (const bus of buses) {
        const persons = state.personsByBus[bus] || [];
        if (persons.length === 0) continue;

        // Get bus-specific settings
        const busSetting = state.busSetting[bus] || state.busSetting['A'];
        const footerText = busSetting.footerText;
        const bgColorHex = busSetting.bgColor.replace('#', '');
        const textColorHex = busSetting.textColor.replace('#', '');
        const borderColorHex = busSetting.borderColor.replace('#', '');

        const totalPages = Math.ceil(persons.length / cellsPerPage);

        for (let pageIndex = 0; pageIndex < totalPages; pageIndex++) {
            const tableRows = [];

            for (let rowIdx = 0; rowIdx < rows; rowIdx++) {
                const cells = [];

                for (let colIdx = 0; colIdx < cols; colIdx++) {
                    const personIndex = pageIndex * cellsPerPage + rowIdx * cols + colIdx;

                    if (personIndex < persons.length) {
                        const person = persons[personIndex];
                        const isVeg = person.isVeg || (person.note && person.note.includes('素'));
                        const nameFontSize = person.override.fontSize || state.fontSizes.name;

                        cells.push(
                            new TableCell({
                                width: { size: Math.floor(100 / cols), type: WidthType.PERCENTAGE },
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        spacing: { after: 100 },
                                        children: [
                                            new TextRun({ text: `● ${state.config.companyName} ●`, size: mmToHalfPoints(state.fontSizes.company), color: borderColorHex })
                                        ]
                                    }),
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        spacing: { after: 100 },
                                        children: [
                                            new TextRun({ text: state.config.eventName, size: mmToHalfPoints(state.fontSizes.event), bold: true, color: textColorHex })
                                        ]
                                    }),
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        spacing: { after: 200 },
                                        children: [
                                            new TextRun({ text: state.config.travelInfo, size: mmToHalfPoints(state.fontSizes.travel), color: borderColorHex })
                                        ]
                                    }),
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        spacing: { before: 200, after: 200 },
                                        children: [
                                            new TextRun({ text: person.name, size: mmToHalfPoints(nameFontSize), bold: true, color: textColorHex })
                                        ]
                                    }),
                                    new Paragraph({
                                        spacing: { after: 80 },
                                        children: [
                                            new TextRun({ text: `${state.config.busLabel}`, size: mmToHalfPoints(state.fontSizes.labels), color: borderColorHex }),
                                            new TextRun({ text: `${person.bus} 車`, size: mmToHalfPoints(state.fontSizes.labels), bold: true, color: textColorHex })
                                        ]
                                    }),
                                    new Paragraph({
                                        spacing: { after: 80 },
                                        children: [
                                            new TextRun({ text: `${state.config.tableLabel}`, size: mmToHalfPoints(state.fontSizes.labels), color: borderColorHex }),
                                            new TextRun({ text: person.table, size: mmToHalfPoints(state.fontSizes.labels), bold: true, color: textColorHex }),
                                            ...(isVeg ? [new TextRun({ text: " 素食桌", size: mmToHalfPoints(state.fontSizes.labels * 0.9), color: "2a6a2a" })] : [])
                                        ]
                                    }),
                                    new Paragraph({
                                        spacing: { after: 100 },
                                        children: [
                                            new TextRun({ text: `${state.config.roomLabel}`, size: mmToHalfPoints(state.fontSizes.labels), color: borderColorHex }),
                                            new TextRun({ text: person.room, size: mmToHalfPoints(state.fontSizes.labels), bold: true, color: textColorHex })
                                        ]
                                    }),
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        spacing: { before: 100 },
                                        children: [
                                            new TextRun({ text: footerText, size: mmToHalfPoints(state.fontSizes.footer), color: borderColorHex })
                                        ]
                                    })
                                ],
                                shading: { fill: bgColorHex },
                                margins: { top: 200, bottom: 200, left: 200, right: 200 }
                            })
                        );
                    } else {
                        // Empty cell
                        cells.push(
                            new TableCell({
                                width: { size: Math.floor(100 / cols), type: WidthType.PERCENTAGE },
                                children: [new Paragraph({ text: "" })],
                                shading: { fill: "FFFFFF" }
                            })
                        );
                    }
                }

                tableRows.push(new TableRow({ children: cells }));
            }

            const table = new Table({
                rows: tableRows,
                width: { size: 100, type: WidthType.PERCENTAGE }
            });

            sections.push({
                properties: {
                    page: {
                        size: {
                            orientation: PageOrientation.PORTRAIT,
                            width: 11906, // A4 width in twips
                            height: 16838 // A4 height in twips
                        },
                        margin: { top: 0, right: 0, bottom: 0, left: 0 }
                    }
                },
                children: [table]
            });
        }
    }

    return new Document({ sections });
}

function setupExportListeners() {
    elements.exportPdf.addEventListener('click', exportPdf);
    elements.exportSelectedPdf.addEventListener('click', exportSelectedPdf);
    elements.exportWord.addEventListener('click', exportWord);
    elements.exportSelectedWord.addEventListener('click', exportSelectedWord);
    elements.printBtn.addEventListener('click', printNametags);

    // Select all / Deselect all buttons
    elements.selectAllBuses.addEventListener('click', () => {
        elements.busCheckboxes.querySelectorAll('input[type="checkbox"]').forEach(cb => {
            cb.checked = true;
        });
    });

    elements.deselectAllBuses.addEventListener('click', () => {
        elements.busCheckboxes.querySelectorAll('input[type="checkbox"]').forEach(cb => {
            cb.checked = false;
        });
    });
}

// ===== Initialize =====
function init() {
    document.getElementById('clearDataBtn').addEventListener('click', clearData);

    setupUploadZone();
    setupConfigListeners();
    setupModalListeners();
    setupPanelToggles();
    setupExportListeners();
}

// Start the app when DOM is ready
document.addEventListener('DOMContentLoaded', init);
