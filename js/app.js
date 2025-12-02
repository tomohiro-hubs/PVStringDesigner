/**
 * PV Module Calculation Logic
 * Based on user provided PDF requirements and standard IEC formulas
 */

// Configuration and Default Data
const CONFIG = {
    defaults: {
        model_name: 'CHSM78N(DG)/F-BH-635',
        pmax_stc: 635,
        voc_stc: 56.41,
        vmp_stc: 46.79,
        isc_stc: 14.35,
        imp_stc: 13.68,
        alpha: 0.043,
        beta: -0.25,
        gamma: -0.29,
        n_start: 14,
        n_end: 19,
        temps: [-20, -15, -10, -5, 0, 10, 20, 25, 48.9, 49.5, 60.9, 60, 70, 80]
    },
    dom: {
        inputs: {
            maker: 'maker_name',
            modelSelect: 'model_select',
            modelCustom: 'model_name_custom',
            pmax: 'pmax_stc',
            voc: 'voc_stc',
            vmp: 'vmp_stc',
            isc: 'isc_stc',
            imp: 'imp_stc',
            alpha: 'alpha',
            beta: 'beta',
            gamma: 'gamma',
            nStart: 'n_series_start',
            nEnd: 'n_series_end',
            maxSysVoltage: 'max_system_voltage',
            // Simple Check Tool
            checkVolt: 'check_voltage',
            checkTolPlus: 'check_tolerance_plus',
            checkTolMinus: 'check_tolerance_minus',
            resMin: 'res_min_voltage',
            resMax: 'res_max_voltage'
        },
        buttons: {
            calculate: 'btnCalculate',
            clear: 'btnClear',
            clearInputs: 'btnClearInputs',
            add: 'btnAddRow',
            export: 'btnExport'
        },
        toggle: {
            mode: 'modeToggle',
            label: 'modeLabel'
        },
        table: {
            headerTop: 'headerRowTop',
            headerBottom: 'headerRowBottom',
            body: 'tableBody',
            template: 'rowTemplate'
        }
    }
};

// App State
let state = {
    mode: 'standard', // 'standard' (beta for voltage) or 'compatibility' (gamma for voltage)
    currentSeriesRange: [],
    moduleData: [] // Loaded from JSON
};

/**
 * Initialize Application
 */
document.addEventListener('DOMContentLoaded', async () => {
    await loadModuleData();
    initializeInputs();
    setupEventListeners();
    // Pre-populate table with default temperatures
    CONFIG.defaults.temps.forEach(temp => addRow(temp));
    // Initial calculation
    calculateAll();
    calculateTolerance(); // Init simple check tool
});

/**
 * Load Module Data from JSON
 */
async function loadModuleData() {
    try {
        const response = await fetch('js/data.json');
        state.moduleData = await response.json();
        populateModels();
    } catch (error) {
        console.error('Failed to load module data:', error);
    }
}

/**
 * Populate Model Select Box with Optgroups by Maker
 */
function populateModels() {
    const modelSelect = document.getElementById(CONFIG.dom.inputs.modelSelect);
    modelSelect.innerHTML = '<option value="">型式を選択してください</option>';
    
    // Group data by maker
    const grouped = state.moduleData.reduce((acc, item) => {
        if (!acc[item.maker]) acc[item.maker] = [];
        acc[item.maker].push(item);
        return acc;
    }, {});

    // Sort makers alphabetically
    const makers = Object.keys(grouped).sort();

    makers.forEach(maker => {
        const group = document.createElement('optgroup');
        group.label = maker;
        
        // Sort models within maker
        grouped[maker].sort((a, b) => a.model.localeCompare(b.model));
        
        grouped[maker].forEach(item => {
            const option = document.createElement('option');
            option.value = item.model;
            option.textContent = item.model;
            group.appendChild(option);
        });
        
        modelSelect.appendChild(group);
    });

    // Add Custom Option at the very end
    const customOption = document.createElement('option');
    customOption.value = "custom";
    customOption.textContent = "カスタム (手入力)";
    modelSelect.appendChild(customOption);
    
    modelSelect.disabled = false;
}

/**
 * Populate Maker Select Box
 */
// function populateMakers() { ... } // Removed

/**
 * Populate Model Select Box based on Maker
 */
// function updateModelSelect(maker) { ... } // Removed

/**
 * Set up default values in input fields
 */
function initializeInputs() {
    // document.getElementById(CONFIG.dom.inputs.modelName).value = CONFIG.defaults.model_name; // Replaced by select logic
    document.getElementById(CONFIG.dom.inputs.pmax).value = CONFIG.defaults.pmax_stc;
    document.getElementById(CONFIG.dom.inputs.voc).value = CONFIG.defaults.voc_stc;
    document.getElementById(CONFIG.dom.inputs.vmp).value = CONFIG.defaults.vmp_stc;
    document.getElementById(CONFIG.dom.inputs.isc).value = CONFIG.defaults.isc_stc;
    document.getElementById(CONFIG.dom.inputs.imp).value = CONFIG.defaults.imp_stc;
    document.getElementById(CONFIG.dom.inputs.alpha).value = CONFIG.defaults.alpha;
    document.getElementById(CONFIG.dom.inputs.beta).value = CONFIG.defaults.beta;
    document.getElementById(CONFIG.dom.inputs.gamma).value = CONFIG.defaults.gamma;
    document.getElementById(CONFIG.dom.inputs.nStart).value = CONFIG.defaults.n_start;
    document.getElementById(CONFIG.dom.inputs.nEnd).value = CONFIG.defaults.n_end;
    
    // Set toggle state (default unchecked = standard)
    const toggle = document.getElementById(CONFIG.dom.toggle.mode);
    toggle.checked = false;
    updateModeLabel();
}

/**
 * Setup Event Listeners
 */
function setupEventListeners() {
    // Buttons
    document.getElementById(CONFIG.dom.buttons.calculate).addEventListener('click', calculateAll);
    document.getElementById(CONFIG.dom.buttons.clear).addEventListener('click', clearResults);
    document.getElementById(CONFIG.dom.buttons.clearInputs).addEventListener('click', clearInputs);
    document.getElementById(CONFIG.dom.buttons.add).addEventListener('click', () => addRow(25));
    document.getElementById(CONFIG.dom.buttons.export).addEventListener('click', exportToExcel);

    // Maker Change - Removed
    /*
    document.getElementById(CONFIG.dom.inputs.maker).addEventListener('change', (e) => {
        updateModelSelect(e.target.value);
        document.getElementById(CONFIG.dom.inputs.modelCustom).classList.add('hidden');
    });
    */

    // Model Change
    document.getElementById(CONFIG.dom.inputs.modelSelect).addEventListener('change', (e) => {
        const model = e.target.value;
        const customInput = document.getElementById(CONFIG.dom.inputs.modelCustom);
        const makerInput = document.getElementById(CONFIG.dom.inputs.maker);
        
        if (model === 'custom') {
            customInput.classList.remove('hidden');
            makerInput.readOnly = false; // Allow manual edit
            makerInput.value = ''; // Clear maker
            makerInput.classList.remove('bg-gray-100');
            makerInput.classList.add('bg-white');
            clearSpecs(); // Clear for manual input
        } else {
            customInput.classList.add('hidden');
            makerInput.readOnly = true;
            makerInput.classList.add('bg-gray-100');
            makerInput.classList.remove('bg-white');
            
            if (model) {
                const data = state.moduleData.find(item => item.model === model);
                if (data) {
                    fillSpecs(data);
                    makerInput.value = data.maker; // Auto-fill maker
                }
            } else {
                makerInput.value = '';
                clearSpecs();
            }
        }
    });

    // Toggle
    const toggle = document.getElementById(CONFIG.dom.toggle.mode);
    toggle.addEventListener('change', (e) => {
        state.mode = e.target.checked ? 'compatibility' : 'standard';
        updateModeLabel();
        calculateAll(); // Re-calculate immediately on mode switch
    });

    // Table delegation for delete buttons
    document.getElementById(CONFIG.dom.table.body).addEventListener('click', (e) => {
        const btn = e.target.closest('.btn-delete-row');
        if (btn) {
            btn.closest('tr').remove();
        }
    });

    // Tolerance Check Tool (Realtime)
    const checkInputs = [CONFIG.dom.inputs.checkVolt, CONFIG.dom.inputs.checkTolPlus, CONFIG.dom.inputs.checkTolMinus];
    checkInputs.forEach(id => {
        document.getElementById(id).addEventListener('input', calculateTolerance);
    });
}

/**
 * Update the text label for the mode toggle
 */
function updateModeLabel() {
    const label = document.getElementById(CONFIG.dom.toggle.label);
    if (state.mode === 'standard') {
        label.textContent = '標準 (βで電圧補正)';
        label.classList.remove('text-blue-600');
        label.classList.add('text-gray-700');
    } else {
        label.textContent = '互換 (γで電圧補正)';
        label.classList.add('text-blue-600');
        label.classList.remove('text-gray-700');
    }
}

/**
 * Add a new row to the table
 */
function addRow(temperature = 25) {
    const template = document.getElementById(CONFIG.dom.table.template);
    const tbody = document.getElementById(CONFIG.dom.table.body);
    const clone = template.content.cloneNode(true);
    
    const input = clone.querySelector('.input-temp');
    input.value = temperature;
    
    // Allow hitting Enter in temp input to trigger calculation
    input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') calculateAll();
    });
    
    // Initialize series cells based on current range if already calculated, 
    // otherwise calculateAll will handle it.
    // But to avoid empty looking row before calc, we rely on calculateAll usually.
    
    tbody.appendChild(clone);
}

/**
 * Clear calculated values (keep temperatures)
 */
function clearResults() {
    // Since we dynamically generate columns, clearing essentially means resetting text
    // but we might want to keep the columns structure.
    const rows = document.querySelectorAll(`#${CONFIG.dom.table.body} tr`);
    rows.forEach(row => {
        row.querySelector('.cell-voc').textContent = '-';
        row.querySelector('.cell-vmp').textContent = '-';
        row.querySelector('.cell-pmax').textContent = '-';
        row.querySelector('.cell-ratio').textContent = '-';
        
        // Clear dynamic series cells
        const dynCells = row.querySelectorAll('.cell-series-val');
        dynCells.forEach(cell => {
            cell.innerHTML = '-';
            cell.classList.remove('text-warning');
        });
    });
}

/**
 * Clear all input fields
 */
function clearInputs() {
    if (!confirm('入力値をすべて消去しますか？')) return;

    // Text inputs
    document.getElementById(CONFIG.dom.inputs.maker).value = '';
    document.getElementById(CONFIG.dom.inputs.modelSelect).value = ''; // Reset Select
    // document.getElementById(CONFIG.dom.inputs.modelSelect).disabled = true; // No longer disabled
    document.getElementById(CONFIG.dom.inputs.modelCustom).value = '';
    document.getElementById(CONFIG.dom.inputs.modelCustom).classList.add('hidden');
    
    // Reset Maker input state to default
    const makerInput = document.getElementById(CONFIG.dom.inputs.maker);
    makerInput.readOnly = true;
    makerInput.classList.add('bg-gray-100');
    makerInput.classList.remove('bg-white');
    
    // Number inputs
    document.getElementById(CONFIG.dom.inputs.pmax).value = '';
    document.getElementById(CONFIG.dom.inputs.voc).value = '';
    document.getElementById(CONFIG.dom.inputs.vmp).value = '';
    document.getElementById(CONFIG.dom.inputs.isc).value = '';
    document.getElementById(CONFIG.dom.inputs.imp).value = '';
    
    // Coeffs
    document.getElementById(CONFIG.dom.inputs.alpha).value = '';
    document.getElementById(CONFIG.dom.inputs.beta).value = '';
    document.getElementById(CONFIG.dom.inputs.gamma).value = '';
    
    // System Config
    document.getElementById(CONFIG.dom.inputs.nStart).value = '';
    document.getElementById(CONFIG.dom.inputs.nEnd).value = '';
    // Max voltage usually stays at standard 1500 or 1000, but clearing it too as requested
    document.getElementById(CONFIG.dom.inputs.maxSysVoltage).value = '';

    // Simple Check Tool - Reset to defaults or clear?
    document.getElementById(CONFIG.dom.inputs.checkVolt).value = '';
    document.getElementById(CONFIG.dom.inputs.checkTolPlus).value = '';
    document.getElementById(CONFIG.dom.inputs.checkTolMinus).value = '';
    calculateTolerance();
}

/**
 * Simple Tolerance Calculation
 */
function calculateTolerance() {
    const v = parseFloat(document.getElementById(CONFIG.dom.inputs.checkVolt).value);
    const tolPlus = parseFloat(document.getElementById(CONFIG.dom.inputs.checkTolPlus).value);
    const tolMinus = parseFloat(document.getElementById(CONFIG.dom.inputs.checkTolMinus).value);

    const elMin = document.getElementById(CONFIG.dom.inputs.resMin);
    const elMax = document.getElementById(CONFIG.dom.inputs.resMax);

    if (isNaN(v) || isNaN(tolPlus) || isNaN(tolMinus)) {
        elMin.textContent = '-';
        elMax.textContent = '-';
        return;
    }

    // Min = V * (100 - tolMinus) / 100
    const minVal = v * ((100 - tolMinus) / 100);
    // Max = V * (100 + tolPlus) / 100
    const maxVal = v * ((100 + tolPlus) / 100);

    // Format: max 2 decimals, remove trailing zeros if integer
    const fmt = (n) => parseFloat(n.toFixed(2)).toLocaleString('ja-JP');

    elMin.textContent = fmt(minVal);
    elMax.textContent = fmt(maxVal);
}

/**
 * Fill spec inputs with data
 */
function fillSpecs(data) {
    document.getElementById(CONFIG.dom.inputs.pmax).value = data.pmax;
    document.getElementById(CONFIG.dom.inputs.voc).value = data.voc;
    document.getElementById(CONFIG.dom.inputs.vmp).value = data.vmp;
    document.getElementById(CONFIG.dom.inputs.isc).value = data.isc;
    document.getElementById(CONFIG.dom.inputs.imp).value = data.imp;
    document.getElementById(CONFIG.dom.inputs.beta).value = data.beta; // Voc coeff
    document.getElementById(CONFIG.dom.inputs.alpha).value = data.alpha; // Isc coeff
    // Gamma from data
    document.getElementById(CONFIG.dom.inputs.gamma).value = data.gamma || ''; 
}

/**
 * Clear spec inputs
 */
function clearSpecs() {
    document.getElementById(CONFIG.dom.inputs.pmax).value = '';
    document.getElementById(CONFIG.dom.inputs.voc).value = '';
    document.getElementById(CONFIG.dom.inputs.vmp).value = '';
    document.getElementById(CONFIG.dom.inputs.isc).value = '';
    document.getElementById(CONFIG.dom.inputs.imp).value = '';
    document.getElementById(CONFIG.dom.inputs.beta).value = '';
    document.getElementById(CONFIG.dom.inputs.alpha).value = '';
    document.getElementById(CONFIG.dom.inputs.gamma).value = '';
}

/**
 * Rebuild Table Header based on Series Range
 */
function rebuildHeader(startN, endN) {
    const headerTop = document.getElementById(CONFIG.dom.table.headerTop);
    const headerBottom = document.getElementById(CONFIG.dom.table.headerBottom);
    
    // Remove existing dynamic headers (anything after the fixed columns)
    // Fixed columns count: 6 in Top (Op, Temp, Voc, Vmp, Pmax, Ratio)
    // But be careful not to remove fixed columns.
    // We marked dynamic headers with class 'dynamic-header'
    
    document.querySelectorAll('.dynamic-header').forEach(el => el.remove());
    
    // Generate new headers
    for (let n = startN; n <= endN; n++) {
        // Top Header: "N直列" spanning 2 columns
        const thTop = document.createElement('th');
        thTop.className = 'px-4 py-3 text-center text-xs font-bold text-gray-700 uppercase tracking-wider bg-gray-100 border-b border-r border-gray-300 dynamic-header';
        thTop.colSpan = 2;
        thTop.innerHTML = `${n}直列`;
        headerTop.appendChild(thTop);
        
        // Bottom Header: "起動(Voc)" and "動作(Vmp)"
        const thVoc = document.createElement('th');
        thVoc.className = 'px-2 py-2 text-right text-[10px] font-medium text-gray-600 bg-gray-50 border-b border-gray-200 dynamic-header w-20';
        thVoc.innerHTML = '起動(Voc)';
        
        const thVmp = document.createElement('th');
        thVmp.className = 'px-2 py-2 text-right text-[10px] font-medium text-gray-600 bg-gray-50 border-b border-r border-gray-200 dynamic-header w-20';
        thVmp.innerHTML = '動作(Vmp)';
        
        headerBottom.appendChild(thVoc);
        headerBottom.appendChild(thVmp);
    }
    
    state.currentSeriesRange = [];
    for(let i=startN; i<=endN; i++) state.currentSeriesRange.push(i);
}

/**
 * Core Calculation Logic
 */
function calculateAll() {
    // Get Input Values
    const inputs = {
        pmax: parseFloat(document.getElementById(CONFIG.dom.inputs.pmax).value) || 0,
        voc: parseFloat(document.getElementById(CONFIG.dom.inputs.voc).value) || 0,
        vmp: parseFloat(document.getElementById(CONFIG.dom.inputs.vmp).value) || 0,
        beta: parseFloat(document.getElementById(CONFIG.dom.inputs.beta).value) || 0,
        gamma: parseFloat(document.getElementById(CONFIG.dom.inputs.gamma).value) || 0,
        nStart: parseInt(document.getElementById(CONFIG.dom.inputs.nStart).value) || 14,
        nEnd: parseInt(document.getElementById(CONFIG.dom.inputs.nEnd).value) || 19,
        maxSysVoltage: parseFloat(document.getElementById(CONFIG.dom.inputs.maxSysVoltage).value) || 1500
    };
    
    // Validate range
    if (inputs.nStart > inputs.nEnd) {
        alert("直列数の開始値は終了値以下にしてください。");
        return;
    }
    if ((inputs.nEnd - inputs.nStart) > 20) {
        if(!confirm("直列数の範囲が広すぎると表が見づらくなる可能性があります。続行しますか？")) return;
    }

    // Rebuild headers if range changed (or always to be safe/simple)
    rebuildHeader(inputs.nStart, inputs.nEnd);

    // Convert percentage coefficients to decimal
    const betaDec = inputs.beta / 100;
    const gammaDec = inputs.gamma / 100;
    const voltageCoeff = state.mode === 'compatibility' ? gammaDec : betaDec;

    // Process each row
    const rows = document.querySelectorAll(`#${CONFIG.dom.table.body} tr`);
    
    rows.forEach(row => {
        const tempInput = row.querySelector('.input-temp');
        const temp = parseFloat(tempInput.value);
        
        // Clear old dynamic cells in this row
        row.querySelectorAll('.dynamic-cell').forEach(el => el.remove());
        
        if (isNaN(temp)) return; // Skip invalid rows but maybe should add empty cells?
        
        // Delta T
        const deltaT = temp - 25;

        // Calculate Module Values
        const vocT = inputs.voc * (1 + voltageCoeff * deltaT);
        const vmpT = inputs.vmp * (1 + voltageCoeff * deltaT);
        const pmaxT = inputs.pmax * (1 + gammaDec * deltaT);
        const ratioT = inputs.pmax ? (pmaxT / inputs.pmax) * 100 : 0;

        // Update Single Module Cells
        updateCell(row, '.cell-voc', vocT);
        updateCell(row, '.cell-vmp', vmpT);
        updateCell(row, '.cell-pmax', pmaxT);
        updateCell(row, '.cell-ratio', ratioT, 2); // % is usually 2 decimals
        
        // Calculate and Append Series Values
        for (let n = inputs.nStart; n <= inputs.nEnd; n++) {
            const vocSys = vocT * n;
            const vmpSys = vmpT * n;
            
            // Create Cells
            const tdVoc = document.createElement('td');
            tdVoc.className = 'px-2 py-2 text-right text-gray-900 text-xs border-b border-gray-100 dynamic-cell cell-series-val bg-gray-50/30';
            
            const tdVmp = document.createElement('td');
            tdVmp.className = 'px-2 py-2 text-right text-gray-900 text-xs border-b border-r border-gray-200 dynamic-cell cell-series-val bg-gray-50/30';

            // Format with commas
            const vocSysFmt = vocSys.toLocaleString('ja-JP', { maximumFractionDigits: 1 });
            const vmpSysFmt = vmpSys.toLocaleString('ja-JP', { maximumFractionDigits: 1 });

            // Warning Check
            if (vocSys > inputs.maxSysVoltage) {
                tdVoc.innerHTML = `<span class="font-bold text-red-600">${vocSysFmt}</span>`;
                tdVoc.title = `Over ${inputs.maxSysVoltage}V`;
            } else {
                tdVoc.textContent = vocSysFmt;
            }
            
            tdVmp.textContent = vmpSysFmt;
            
            row.appendChild(tdVoc);
            row.appendChild(tdVmp);
        }
    });
}

/**
 * Helper to update a cell text
 */
function updateCell(row, selector, value) {
    const cell = row.querySelector(selector);
    // Format with commas, max 2 decimals
    const text = value.toLocaleString('ja-JP', { maximumFractionDigits: 2 }); 
    cell.textContent = text;
    cell.classList.remove('value-updated');
    void cell.offsetWidth;
    cell.classList.add('value-updated');
}

/**
 * Export data to Excel
 */
function exportToExcel() {
    // 1. Gather Input Data
    let modelName = document.getElementById(CONFIG.dom.inputs.modelSelect).value;
    if (modelName === 'custom') {
        modelName = document.getElementById(CONFIG.dom.inputs.modelCustom).value;
    }
    // Fallback if logic is bypassed
    if (!modelName && document.getElementById(CONFIG.dom.inputs.modelCustom).value) {
        modelName = document.getElementById(CONFIG.dom.inputs.modelCustom).value;
    }

    const inputs = {
        maker: document.getElementById(CONFIG.dom.inputs.maker).value,
        modelName: modelName,
        pmax: document.getElementById(CONFIG.dom.inputs.pmax).value,
        voc: document.getElementById(CONFIG.dom.inputs.voc).value,
        vmp: document.getElementById(CONFIG.dom.inputs.vmp).value,
        isc: document.getElementById(CONFIG.dom.inputs.isc).value,
        imp: document.getElementById(CONFIG.dom.inputs.imp).value,
        beta: document.getElementById(CONFIG.dom.inputs.beta).value,
        gamma: document.getElementById(CONFIG.dom.inputs.gamma).value,
        alpha: document.getElementById(CONFIG.dom.inputs.alpha).value,
        nStart: document.getElementById(CONFIG.dom.inputs.nStart).value,
        nEnd: document.getElementById(CONFIG.dom.inputs.nEnd).value,
        mode: state.mode === 'standard' ? '標準 (β使用)' : '互換 (γ使用)'
    };

    // 2. Prepare Data Array for Sheet
    const sheetData = [
        ['PVモジュール温度特性計算結果', '', '', '', ''],
        [''],
        ['【設定パラメータ】'],
        ['項目', '値', '単位'],
        ['メーカー', inputs.maker, ''],
        ['型式', inputs.modelName, ''],
        ['Pmax (STC)', parseFloat(inputs.pmax), 'W'],
        ['Voc (STC)', parseFloat(inputs.voc), 'V'],
        ['Vmp (STC)', parseFloat(inputs.vmp), 'V'],
        ['Isc (STC)', parseFloat(inputs.isc), 'A'],
        ['Imp (STC)', parseFloat(inputs.imp), 'A'],
        ['直列数範囲', `${inputs.nStart} 〜 ${inputs.nEnd}`, '直列'],
        ['温度係数 β (電圧)', parseFloat(inputs.beta), '%/℃'],
        ['温度係数 γ (出力)', parseFloat(inputs.gamma), '%/℃'],
        ['計算モード', inputs.mode, ''],
        [''],
        ['【計算結果】']
    ];

    // Build Header Row
    const headerRow = ['温度 (℃)', 'Voc (V)', 'Vmp (V)', 'Pmax (W)', '比率 (%)'];
    state.currentSeriesRange.forEach(n => {
        headerRow.push(`${n}直列 Voc`);
        headerRow.push(`${n}直列 Vmp`);
    });
    sheetData.push(headerRow);

    // 3. Gather Table Data
    const rows = document.querySelectorAll(`#${CONFIG.dom.table.body} tr`);
    rows.forEach(row => {
        const temp = parseFloat(row.querySelector('.input-temp').value);
        
        // Helper to parse formatted numbers (remove commas)
        const parseVal = (sel) => {
            const text = row.querySelector(sel).textContent.replace(/,/g, '').trim();
            return text === '-' ? NaN : parseFloat(text);
        };

        const voc = parseVal('.cell-voc');
        const vmp = parseVal('.cell-vmp');
        const pmax = parseVal('.cell-pmax');
        const ratio = parseVal('.cell-ratio');
        
        const rowData = [
            temp,
            isNaN(voc) ? '-' : voc,
            isNaN(vmp) ? '-' : vmp,
            isNaN(pmax) ? '-' : pmax,
            isNaN(ratio) ? '-' : ratio
        ];
        
        // Dynamic Cells
        const seriesCells = row.querySelectorAll('.cell-series-val');
        seriesCells.forEach(cell => {
            // Handle HTML inside cell (warning span) and commas
            const text = cell.textContent.replace(/,/g, '').trim();
            const val = parseFloat(text);
            rowData.push(isNaN(val) ? '-' : val);
        });
        
        sheetData.push(rowData);
    });

    // 4. Create Workbook
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(sheetData);

    // Widths
    const wscols = [
        {wch: 15}, // Temp
        {wch: 12}, // Voc
        {wch: 12}, // Vmp
        {wch: 12}, // Pmax
        {wch: 10}  // Ratio
    ];
    // Add dynamic col widths
    state.currentSeriesRange.forEach(() => {
        wscols.push({wch: 12});
        wscols.push({wch: 12});
    });
    ws['!cols'] = wscols;

    XLSX.utils.book_append_sheet(wb, ws, "計算結果");

    // 5. Download
    const dateStr = new Date().toISOString().slice(0,10).replace(/-/g, '');
    XLSX.writeFile(wb, `PV計算結果_${inputs.nStart}-${inputs.nEnd}直列_${dateStr}.xlsx`);
}