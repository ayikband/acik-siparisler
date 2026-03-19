/**
 * Ayık Band Open Orders Dashboard - Main Logic
 */

const state = {
    rawData: [],
    filteredData: [],
    filters: {
        month: '', rep: '', sender: '', sector: '', city: '', search: ''
    },
    settings: {
        rowsPerPage: 25, currentPage: 1, sortCol: 'date', sortAsc: false // En yeni siparişler üstte
    },
    charts: {},
    rates: { USD: 1.08, GBP: 0.85, TRY: 35.0 } // 1 EUR = ? USD/GBP/TRY
};

const COLORS = ['#1F76AC', '#72B2E2', '#27C485', '#F1C40F', '#E74C3C', '#9B59B6', '#16A085', '#34495E', '#D35400', '#7F8C8D'];

// YENİ EXCEL BAŞLIK HARİTASI
// YENİ EXCEL BAŞLIK HARİTASI
const COLS = {
    FIRM: "Müşteri Adı",
    MATERIAL: "Malzeme",
    QTY: "Adet/Miktar",
    CITY: "Şehir",
    DELIVERY_CITY: "Sevk Yeri Şehri", // Yeni Eklendi
    SECTOR: "Sektör",
    REP: "Satış Temsilcisi",
    SENDER: "Teklifi Gönderen",
    ORDER_NO: "Sipariş No",
    ORDER_DATE: "Sipariş Tarihi",
    VAL_EUR: "Toplam Fiyat/EUR",
    VAL_USD: "Toplam Fiyat/USD",
    VAL_GBP: "Toplam Fiyat/GBP",
    VAL_TL: "Toplam Fiyat/TL"
};

const FORMATTER = {
    currency: (val, curr = 'EUR') => new Intl.NumberFormat('tr-TR', { style: 'currency', currency: curr }).format(val),
    number: (val) => new Intl.NumberFormat('tr-TR', { maximumFractionDigits: 2 }).format(val),
    date: (date) => date ? new Date(date).toLocaleDateString('tr-TR') : '-',
    percent: (val, d = 1) => '%' + new Intl.NumberFormat('tr-TR', { maximumFractionDigits: d, minimumFractionDigits: d }).format(val)
};

document.addEventListener('DOMContentLoaded', () => {
    initEvents();
    loadRates();
});

function initEvents() {
    document.getElementById('fileInput').addEventListener('change', handleFileUpload);

    ['filterMonth', 'filterRep', 'filterSender', 'filterSector', 'filterCity'].forEach(id => {
        document.getElementById(id).addEventListener('change', (e) => updateFilter(e.target.id, e.target.value));
    });
    document.getElementById('custSearch').addEventListener('input', (e) => updateFilter('search', e.target.value));
    document.getElementById('resetFiltersBtn').addEventListener('click', resetFilters);

    document.getElementById('ratesBtn').addEventListener('click', openRatesModal);
    document.getElementById('closeRatesBtn').addEventListener('click', () => toggleModal('ratesModal', false));
    document.getElementById('saveRatesBtn').addEventListener('click', saveRates);

    document.getElementById('prevPageFn').addEventListener('click', () => changePage(-1));
    document.getElementById('nextPageFn').addEventListener('click', () => changePage(1));
    document.getElementById('rowsPerPage').addEventListener('change', (e) => {
        state.settings.rowsPerPage = parseInt(e.target.value);
        state.settings.currentPage = 1;
        renderTable();
    });

    document.getElementById('cumulativeToggle').addEventListener('change', renderCharts);
    document.getElementById('weeklyToggle').addEventListener('change', renderCharts);
    // initEvents fonksiyonunun içindeki uygun bir yere (örneğin diğer dinleyicilerin arasına) şunu yapıştır:

    document.getElementById('exportPdfBtn').addEventListener('click', exportDashboardToPDF);
    document.querySelectorAll('#detailTable th[data-sort]').forEach(th => {
        th.addEventListener('click', () => {
            const field = th.dataset.sort;
            if (state.settings.sortCol === field) state.settings.sortAsc = !state.settings.sortAsc;
            else { state.settings.sortCol = field; state.settings.sortAsc = true; }
            sortData(); renderTable();
        });
    });
}

function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    document.getElementById('uploadStatus').textContent = "Dosya okunuyor...";

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            // --- AKILLI BAŞLIK BULUCU BAŞLANGICI ---
            // Önce tüm sayfayı dizi (array) olarak alıp başlık satırını arıyoruz
            const rawArr = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            let headerRowIdx = 0;

            // İlk 20 satırı tara, içinde "Müşteri Adı" geçen satırı başlık olarak kabul et
            for (let i = 0; i < Math.min(20, rawArr.length); i++) {
                if (rawArr[i] && rawArr[i].includes("Müşteri Adı")) {
                    headerRowIdx = i;
                    console.log("Başlıklar şu satırda bulundu:", i + 1);
                    break;
                }
            }

            // Artık SheetJS'e verileri okumaya tam olarak o satırdan (range) başlamasını söylüyoruz
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { range: headerRowIdx, defval: "" });
            // --- AKILLI BAŞLIK BULUCU BİTİŞİ ---

            processRawData(jsonData);

            document.getElementById('uploadOverlay').classList.add('hidden');
            document.getElementById('mainDashboard').classList.remove('hidden');
        } catch (err) {
            console.error(err);
            document.getElementById('uploadStatus').textContent = "Hata: " + err.message;
        }
    };
    reader.readAsArrayBuffer(file);
}

function parseNumberTR(val) {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    let clean = val.toString().replace(/\./g, "").replace(",", ".");
    return parseFloat(clean) || 0;
}

function parseDateTR(dateRaw) {
    if (!dateRaw) return null;
    if (typeof dateRaw === 'number') return new Date(Math.round((dateRaw - 25569) * 86400 * 1000));
    if (typeof dateRaw === 'string') {
        const parts = dateRaw.split('.');
        if (parts.length === 3) return new Date(parts[2], parts[1] - 1, parts[0]);
    }
    return null;
}

function processRawData(json) {
    state.rawData = json.map((row, idx) => {
        return {
            id: idx,
            firm: (row[COLS.FIRM] || "").trim(),
            date: parseDateTR(row[COLS.ORDER_DATE]),
            orderNo: row[COLS.ORDER_NO],
            material: row[COLS.MATERIAL],
            city: (row[COLS.CITY] || "Belirtilmedi").trim(),
            deliveryCity: (row[COLS.DELIVERY_CITY] || "Belirtilmedi").trim(), // Yeni Eklendi
            sector: (row[COLS.SECTOR] || "Diğer").trim(),
            rep: (row[COLS.REP] || "Belirtilmedi").trim(),
            sender: (row[COLS.SENDER] || "").trim(),
            // Para Birimleri (Ayrı Ayrı Saklanıyor)
            valEur: parseNumberTR(row[COLS.VAL_EUR]),
            valUsd: parseNumberTR(row[COLS.VAL_USD]),
            valGbp: parseNumberTR(row[COLS.VAL_GBP]),
            valTl: parseNumberTR(row[COLS.VAL_TL])
        };
    }).filter(r => r.date && r.firm !== ""); // Geçerli tarih ve firma olanları al

    // Dropdownları doldur
    populateDropdown('filterRep', [...new Set(state.rawData.map(r => r.rep))].sort());
    populateDropdown('filterSender', [...new Set(state.rawData.map(r => r.sender))].sort());
    populateDropdown('filterSector', [...new Set(state.rawData.map(r => r.sector))].sort());
    populateDropdown('filterCity', [...new Set(state.rawData.map(r => r.city))].sort());

    const months = [...new Set(state.rawData.map(r => r.date.toLocaleString('tr-TR', { month: 'long', year: 'numeric' })))];
    populateDropdown('filterMonth', months);

    applyFilters();
}

function populateDropdown(id, items) {
    const sel = document.getElementById(id);
    if (!sel) return;
    sel.innerHTML = sel.options[0].outerHTML;
    items.forEach(i => {
        if (!i) return;
        const opt = document.createElement('option');
        opt.value = i;
        opt.textContent = i;
        sel.appendChild(opt);
    });
}

function updateFilter(key, val) {
    const map = { filterMonth: 'month', filterRep: 'rep', filterSender: 'sender', filterSector: 'sector', filterCity: 'city' };
    if (map[key]) state.filters[map[key]] = val;
    if (key === 'search') state.filters.search = val.toLowerCase();

    state.settings.currentPage = 1;
    applyFilters();
}

function resetFilters() {
    state.filters = { month: '', rep: '', sender: '', sector: '', city: '', search: '' };
    document.querySelectorAll('.filter-bar select').forEach(s => s.value = "");
    document.querySelector('.filter-bar input[type="text"]').value = "";
    applyFilters();
}

function applyFilters() {
    state.filteredData = state.rawData.filter(row => {
        const rowMonth = row.date.toLocaleString('tr-TR', { month: 'long', year: 'numeric' });
        return (!state.filters.month || rowMonth === state.filters.month) &&
            (!state.filters.rep || row.rep === state.filters.rep) &&
            (!state.filters.sender || row.sender === state.filters.sender) &&
            (!state.filters.sector || row.sector === state.filters.sector) &&
            (!state.filters.city || row.city === state.filters.city) &&
            (!state.filters.search || row.firm.toLowerCase().includes(state.filters.search));
    });
    updateDashboard();
}

// ANA MANTIK: Orijinal birimleri topla, en son güncel kurla EUR'ya çevir
function getCalculatedEurValue(row) {
    const r = state.rates;
    return row.valEur + (row.valUsd / r.USD) + (row.valGbp / r.GBP) + (row.valTl / r.TRY);
}

function updateDashboard() {
    renderKPIs();
    renderCharts();
    renderRankings();
    sortData();
    renderTable();
}

function renderKPIs() {
    const d = state.filteredData;
    let sumEur = 0, sumUsd = 0, sumGbp = 0, sumTl = 0;

    d.forEach(r => {
        sumEur += r.valEur;
        sumUsd += r.valUsd;
        sumGbp += r.valGbp;
        sumTl += r.valTl;
    });

    // Güncel Kur ile Genel Toplam Hesaplama
    const r = state.rates;
    const grandTotalEur = sumEur + (sumUsd / r.USD) + (sumGbp / r.GBP) + (sumTl / r.TRY);

    document.getElementById('kpiTotalEur').textContent = FORMATTER.currency(grandTotalEur, 'EUR');

    // Orijinal Kırılımlar
    const rowsHTML = [];
    if (sumEur > 0) rowsHTML.push(`<div class="currency-row"><span class="curr-label">EUR</span><span class="curr-value">${FORMATTER.currency(sumEur, 'EUR')}</span></div>`);
    if (sumUsd > 0) rowsHTML.push(`<div class="currency-row"><span class="curr-label">USD</span><span class="curr-value">${FORMATTER.currency(sumUsd, 'USD')}</span></div>`);
    if (sumGbp > 0) rowsHTML.push(`<div class="currency-row"><span class="curr-label">GBP</span><span class="curr-value">${FORMATTER.currency(sumGbp, 'GBP')}</span></div>`);
    if (sumTl > 0) rowsHTML.push(`<div class="currency-row"><span class="curr-label">TL</span><span class="curr-value">${FORMATTER.currency(sumTl, 'TRY')}</span></div>`);

    document.getElementById('kpiCurrencyBreakdown').innerHTML = rowsHTML.length ? rowsHTML.join('') : '<p class="text-muted">Kayıt Yok</p>';

    document.getElementById('kpiOrderCount').textContent = d.length;
    document.getElementById('kpiCustomerCount').textContent = `Bekleyen Müşteri: ${new Set(d.map(x => x.firm)).size}`;

    // Sektör Pasta Grafiği
    const sectorMap = {};
    d.forEach(r => {
        const val = getCalculatedEurValue(r);
        sectorMap[r.sector] = (sectorMap[r.sector] || 0) + val;
    });

    const secKeys = Object.keys(sectorMap).sort((a, b) => sectorMap[b] - sectorMap[a]).slice(0, 5); // İlk 5 sektör
    const secData = secKeys.map((k, i) => ({ label: k, value: sectorMap[k], color: COLORS[i % COLORS.length] }));

    const ctxSector = document.getElementById('kpiChartSector').getContext('2d');
    createOrUpdateChart('sector', ctxSector, {
        type: 'doughnut',
        data: {
            labels: secData.map(d => d.label),
            datasets: [{ data: secData.map(d => d.value), backgroundColor: secData.map(d => d.color), borderWidth: 0 }]
        },
        options: { cutout: '70%', maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });
    generateLegend('kpiLegendSector', secData);
}

function renderCharts() {
    Chart.defaults.color = '#4A5568';
    Chart.defaults.borderColor = '#E2E8F0';

    // 1. Trend Grafiği (Sipariş Tarihine Göre)
    const ctxTrend = document.getElementById('chartTrend').getContext('2d');
    const isCumulative = document.getElementById('cumulativeToggle').checked;
    const isWeekly = document.getElementById('weeklyToggle').checked;

    const pivot = {};
    const dateSet = new Set();

    state.filteredData.forEach(r => {
        let key;
        if (isWeekly) {
            const [y, w] = getISOWeekNumber(r.date);
            key = `${y}-W${w.toString().padStart(2, '0')}`;
        } else {
            key = `${r.date.getFullYear()}-${String(r.date.getMonth() + 1).padStart(2, '0')}-${String(r.date.getDate()).padStart(2, '0')}`;
        }
        dateSet.add(key);
        pivot[key] = (pivot[key] || 0) + getCalculatedEurValue(r);
    });

    const sortedDates = [...dateSet].sort();
    let chartData = [];
    let acc = 0;

    sortedDates.forEach(d => {
        if (isCumulative) { acc += pivot[d]; chartData.push(acc); }
        else { chartData.push(pivot[d]); }
    });

    createOrUpdateChart('trend', ctxTrend, {
        type: isCumulative ? 'line' : 'bar',
        data: {
            labels: sortedDates,
            datasets: [{
                label: 'Açık Sipariş Hacmi (EUR)',
                data: chartData,
                backgroundColor: '#1F76AC',
                borderColor: '#1F76AC',
                fill: isCumulative ? { target: 'origin', above: 'rgba(31, 118, 172, 0.1)' } : false,
                tension: 0.3
            }]
        },
        options: { maintainAspectRatio: false, plugins: { legend: { display: false } } }
    });

    // 2. Temsilci Grafiği
    const repMap = {};
    state.filteredData.forEach(r => {
        repMap[r.rep] = (repMap[r.rep] || 0) + getCalculatedEurValue(r);
    });
    const reps = Object.keys(repMap).sort((a, b) => repMap[b] - repMap[a]);

    const ctxRep = document.getElementById('chartRep').getContext('2d');
    createOrUpdateChart('rep', ctxRep, {
        type: 'bar',
        data: {
            labels: reps,
            datasets: [{ label: 'Bekleyen Bakiye', data: reps.map(k => repMap[k]), backgroundColor: '#27C485' }]
        },
        options: { indexAxis: 'y', maintainAspectRatio: false }
    });

    // 3. Şehir Grafiği (Sevk Yeri Şehri Bazlı)
    const cityMap = {};
    state.filteredData.forEach(r => {
        cityMap[r.deliveryCity] = (cityMap[r.deliveryCity] || 0) + getCalculatedEurValue(r);
    });
    const cities = Object.keys(cityMap).sort((a, b) => cityMap[b] - cityMap[a]).slice(0, 10);

    const ctxCity = document.getElementById('chartCity').getContext('2d');
    createOrUpdateChart('city', ctxCity, {
        type: 'bar',
        data: {
            labels: cities,
            datasets: [{ label: 'Hacim', data: cities.map(k => cityMap[k]), backgroundColor: '#F1C40F' }]
        },
        options: { maintainAspectRatio: false }
    });
}

function sortData() {
    const { sortCol, sortAsc } = state.settings;
    state.filteredData.sort((a, b) => {
        let valA = a[sortCol]; let valB = b[sortCol];
        if (sortCol === 'netEur') { valA = getCalculatedEurValue(a); valB = getCalculatedEurValue(b); }
        if (sortCol === 'customer') { valA = a.firm; valB = b.firm; }
        if (sortCol === 'orderNo') { valA = a.orderNo; valB = b.orderNo; }

        if (valA < valB) return sortAsc ? -1 : 1;
        if (valA > valB) return sortAsc ? 1 : -1;
        return 0;
    });
}

function renderTable() {
    const tbody = document.querySelector('#detailTable tbody');
    tbody.innerHTML = '';
    const { currentPage, rowsPerPage } = state.settings;
    const start = (currentPage - 1) * rowsPerPage;
    const end = start + rowsPerPage;
    const pageData = state.filteredData.slice(start, end);

    pageData.forEach(row => {
        const eurEqv = getCalculatedEurValue(row);

        // Orijinal birimleri bulup text yapalım
        let origText = [];
        if (row.valEur > 0) origText.push(`${FORMATTER.number(row.valEur)} EUR`);
        if (row.valUsd > 0) origText.push(`${FORMATTER.number(row.valUsd)} USD`);
        if (row.valGbp > 0) origText.push(`${FORMATTER.number(row.valGbp)} GBP`);
        if (row.valTl > 0) origText.push(`${FORMATTER.number(row.valTl)} TL`);

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${FORMATTER.date(row.date)}</td>
            <td>${row.orderNo}</td>
            <td title="${row.firm}"><strong>${row.firm.length > 25 ? row.firm.substring(0, 25) + '...' : row.firm}</strong></td>
            <td title="${row.material}"><span style="font-size: 0.8rem; color: #718096;">${row.material.substring(0, 30)}...</span></td>
            <td>${row.rep}</td>
            <td>${row.sector}</td>
            <td class="text-right text-muted">${origText.join('<br>')}</td>
            <td class="text-right text-accent" style="font-weight: 700;">${FORMATTER.number(eurEqv)} €</td>
        `;
        tbody.appendChild(tr);
    });

    document.getElementById('pageInfo').textContent = `${start + 1}-${Math.min(end, state.filteredData.length)} / ${state.filteredData.length}`;
}

function changePage(delta) {
    const total = Math.ceil(state.filteredData.length / state.settings.rowsPerPage);
    const newP = state.settings.currentPage + delta;
    if (newP >= 1 && newP <= total) { state.settings.currentPage = newP; renderTable(); }
}

function loadRates() {
    const saved = localStorage.getItem('ayik_rates_open');
    if (saved) state.rates = JSON.parse(saved);
}

function openRatesModal() {
    document.getElementById('rateUSD').value = state.rates.USD;
    document.getElementById('rateGBP').value = state.rates.GBP;
    document.getElementById('rateTRY').value = state.rates.TRY;
    toggleModal('ratesModal', true);
}

function saveRates() {
    state.rates.USD = parseFloat(document.getElementById('rateUSD').value) || 1.08;
    state.rates.GBP = parseFloat(document.getElementById('rateGBP').value) || 0.85;
    state.rates.TRY = parseFloat(document.getElementById('rateTRY').value) || 35.0;

    localStorage.setItem('ayik_rates_open', JSON.stringify(state.rates));
    toggleModal('ratesModal', false);

    // Kur değiştiği an her şey yeniden hesaplanıp arayüze basılır
    updateDashboard();
}

function toggleModal(id, show) {
    const el = document.getElementById(id);
    if (show) el.classList.remove('hidden'); else el.classList.add('hidden');
}

function createOrUpdateChart(key, ctx, config) {
    if (state.charts[key]) state.charts[key].destroy();
    state.charts[key] = new Chart(ctx, config);
}

function generateLegend(containerId, items) {
    const container = document.getElementById(containerId);
    if (!container) return;
    container.innerHTML = '';
    const total = items.reduce((sum, item) => sum + item.value, 0);

    items.forEach(item => {
        const percent = total > 0 ? Math.round((item.value / total) * 100) : 0;
        container.innerHTML += `
            <div class="legend-item">
                <span class="legend-dot" style="background-color: ${item.color}"></span>
                <span class="legend-label" title="${item.label}">${item.label}</span>
                <span class="legend-val">${FORMATTER.percent(percent, 0)}</span>
            </div>`;
    });
}

function renderRankings() {
    const data = state.filteredData;

    // --- 1. Top customers by total EUR value ---
    const customerMap = {};
    data.forEach(r => {
        const val = getCalculatedEurValue(r);
        if (!customerMap[r.firm]) customerMap[r.firm] = 0;
        customerMap[r.firm] += val;
    });
    const topCustomers = Object.entries(customerMap)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);

    const custContainer = document.getElementById('topCustomersList');
    custContainer.innerHTML = '';
    const maxCust = topCustomers[0]?.[1] || 1;
    topCustomers.forEach(([firm, total], i) => {
        const pct = (total / maxCust) * 100;
        const rankClass = i === 0 ? 'gold' : i === 1 ? 'silver' : i === 2 ? 'bronze' : '';
        const row = document.createElement('div');
        row.className = 'ranking-row';
        row.innerHTML = `
            <div class="ranking-row-fill" style="width:${pct.toFixed(1)}%"></div>
            <span class="ranking-rank ${rankClass}">${i + 1}</span>
            <span class="ranking-name" title="${firm}">${firm}</span>
            <span class="ranking-amount">${FORMATTER.currency(total, 'EUR')}</span>
        `;
        custContainer.appendChild(row);
    });

    // --- 2. Top single orders by EUR equivalent ---
    const topOrders = [...data]
        .sort((a, b) => getCalculatedEurValue(b) - getCalculatedEurValue(a))
        .slice(0, 10);

    const ordersContainer = document.getElementById('topSingleOrdersList');
    ordersContainer.innerHTML = '';
    const maxOrder = topOrders[0] ? getCalculatedEurValue(topOrders[0]) : 1;
    topOrders.forEach((row, i) => {
        const val = getCalculatedEurValue(row);
        const pct = (val / maxOrder) * 100;
        const rankClass = i === 0 ? 'gold' : i === 1 ? 'silver' : i === 2 ? 'bronze' : '';
        const el = document.createElement('div');
        el.className = 'ranking-row';
        el.innerHTML = `
            <div class="ranking-row-fill" style="width:${pct.toFixed(1)}%"></div>
            <span class="ranking-rank ${rankClass}">${i + 1}</span>
            <span class="ranking-name" title="${row.firm}">${row.firm}</span>
            <span class="ranking-sub">${row.orderNo ? '#' + row.orderNo : ''}</span>
            <span class="ranking-amount">${FORMATTER.currency(val, 'EUR')}</span>
        `;
        ordersContainer.appendChild(el);
    });
}

function getISOWeekNumber(d) {
    d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    return [d.getUTCFullYear(), weekNo];
}

// --- PDF ÇIKTI FONKSİYONU ---
async function exportDashboardToPDF() {
    const btn = document.getElementById('exportPdfBtn');
    const originalText = btn.innerHTML;
    btn.innerHTML = '⏳ Hazırlanıyor...';
    btn.disabled = true;

    try {
        // 1. PDF'te görünmemesi gereken bölümleri gizle (Tablo ve Filtreler)
        const filterBar = document.querySelector('.filter-bar');
        const tableControls = document.querySelector('.pagination-controls');
        const tableSection = document.querySelector('.table-section'); // Tabloyu bul

        if (filterBar) filterBar.style.display = 'none';
        if (tableControls) tableControls.style.display = 'none';
        if (tableSection) tableSection.style.display = 'none'; // Tabloyu gizle

        // Expand ranking lists so nothing is clipped by max-height/overflow
        const rankingLists = document.querySelectorAll('.ranking-list');
        rankingLists.forEach(el => { el.style.maxHeight = 'none'; el.style.overflow = 'visible'; });

        // 2. Özel Kurumsal PDF Başlığı Oluşturma
        const printHeader = document.createElement('div');
        printHeader.id = 'pdf-print-header';

        const today = new Date().toLocaleDateString('tr-TR');
        const totalVal = document.getElementById('kpiTotalEur').textContent;
        const r = state.rates; // Kurları al

        // Ayık Band logosu sol üste, detaylar ve kurlar sağ üste
        printHeader.innerHTML = `
            <table width="100%" style="border-bottom: 3px solid #1F76AC; padding-bottom: 15px; margin-bottom: 25px; border-collapse: collapse;">
                <tr>
                    <td style="vertical-align: top; width: 40%;">
                        <img src="data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTM0MCIgaGVpZ2h0PSIzODUiIHZpZXdCb3g9IjAgMCAxMzQwIDM4NSIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZmlsbC1ydWxlPSJldmVub2RkIiBjbGlwLXJ1bGU9ImV2ZW5vZGQiIGQ9Ik0zMTEuMDQgMjE5LjQzTDIwOC41NSA4NEMxODcuOTYgNTMuOTYgMTUzLjM5IDM0LjIzIDExNC4xOSAzNC4yM0M1MS4xMyAzNC4yMyAwIDg1LjIxIDAgMTQ4LjE5QzAgMjExLjE3IDUxLjEzIDI2Mi4wNyAxMTQuMTkgMjYyLjA3QzE1Mi42MyAyNjIuMDcgMTg2LjY3IDI0My4xIDIwNy4zMyAyMTMuOTdMMTY3LjY3IDE1OS45NUMxNjIuMDUgMTg0LjUzIDE0MC4wMSAyMDIuODkgMTEzLjY1IDIwMi44OUM4My4wMyAyMDIuODkgNTguMjcgMTc4LjE1IDU4LjI3IDE0Ny41OEM1OC4yNyAxMTcuMDEgODMuMTEgOTIuMzUgMTEzLjY1IDkyLjM1QzEzMS40MyA5Mi4zNSAxNDcuMjMgMTAwLjY5IDE1Ny40MSAxMTMuNzVMMjE4LjY0IDE5NC4xN0wyNjQuMDcgMjU0LjQxSDI2NC4zOEMyNjkuNjIgMjU5LjExIDI3Ni42MSAyNjIgMjg0LjIxIDI2MkMzMDAuNjIgMjYyIDMxMy44NCAyNDguNzIgMzEzLjg0IDIzMi40MUMzMTMuODQgMjI3LjcxIDMxMi44NSAyMjMuMjMgMzEwLjk1IDIxOS4yOEwzMTEuMDMzMjE5LjQzSDMxMS4wNFpNNDE0LjUyIDM0LjIzQzM3Ni4wOCAzNC4yMyAzNDIuMDQgNTMuMiAzMjEuMzggODIuMzNMMzYxLjA0IDEzNi4zNUMzNjYuNjYgMTExLjc3IDM4OC42OSA5My40MSA0MTUuMDYgOTMuNDFDNDQ1LjY4IDkzLjQxIDQ3MC40NCAxMTguMTQgNDcwLjQ0IDE0OC43MkM0NzAuNDQgMTc5LjMgNDQ1LjYgMjA0LjAzIDQxNS4wNiAyMDQuMDNDMzk3LjI4IDIwNC4wMyAzODEuNDggMTk1LjY5IDM3MS4zIDE4Mi42M0wzMTAuMDcgMTAyLjIxTDI2NC42NCA0MS45N0MyNjQuNjQgNDEuOTcgMjY0LjQxIDQxLjk3IDI2NC4zNCA0MS45QzI1OS4xIDM3LjE5IDI1Mi4xMSAzNC4zMSAyNDQuNTEgMzQuMzFDMjI4LjE4IDM0LjMxIDIxNC44OCA0Ny41OSAyMTQuODggNjMuOUMyMTQuODggNjguNjEgMjE1Ljg3IDczLjA4IDIxNy43NyA3Ni45NUwzMjAuMjYgMjEyLjM4QzM0MC44NSAyNDIuNDMgMzc1LjQxIDI2Mi4wOCA0MTQuNjIgMjYyLjA4QzQ3Ny42OCAyNjIuMDggNTI4LjggMjExLjA5IDUyOC44IDE0OC4yQzUyOC44IDg1LjMxIDQ3Ny42NyAzNC4yNCA0MTQuNjIgMzQuMjRINDE0LjUyVjM0LjIzWiIgZmlsbD0iIzRBOEJDOSIvPgo8cGF0aCBkPSJNOTEuNDUgMzQ5Ljk2SDc5LjMxTDc0LjU2IDMzNi42M0g0OC4wN0w0My40IDM0OS45NkgzMS42TDU0LjYxIDI4OS41MUg2OC4wMkw5MS40NSAzNDkuOTZaTTUzLjg0IDMyMC4zM0w1MS41NSAzMjYuODdINzEuMDhMNjguNyAzMjAuMzNDNjYuNzIgMzE1LjI5IDY0LjIzIDMwOC4zOSA2MS4yMyAyOTkuNjJDNTguODUgMzA2LjY0IDU2LjM5IDMxMy41NCA1My44NCAzMjAuMzNaIiBmaWxsPSIjMDAxRjM4Ii8+CjxwYXRoIGQ9Ik0xMjIuNDMgMzAwLjM4TDEyOS4wNSAyODkuNTFIMTQxLjdMMTE5LjAzIDMyNi4xVjM0OS45NkgxMDcuNTdWMzI2LjFMODQuNzMgMjg5LjUxSDk3LjcyTDEwNC40MyAzMDAuNDZDMTA5Ljg2IDMwOS41MiAxMTIuODkgMzE0LjYxIDExMy41MSAzMTUuNzRDMTE1LjA0IDMxMy4wMyAxMTguMDEgMzA3LjkgMTIyLjQyIDMwMC4zNyIgZmlsbD0iIzAwMUYzOCIvPgo8cGF0aCBkPSJNMTYwLjM4IDI4OS41MUgxNDguOTJWMzQ5Ljk2SDE2MC4zOFYyODkuNTFaIiBmaWxsPSIjMDAxRjM4Ii8+CjxwYXRoIGQ9Ik0yMTMuNDQgMjg5LjUxSDIyNy43TDIwMi44MyAzMTQuMTNMMjI3LjcgMzQ5Ljk2SDIxNC4wNEwxOTQuNjggMzIxLjc3TDE4NS4xNyAzMzEuMDRWMzQ5Ljk2SDE3My43MVYyODkuNTFIMTg1LjE3VjMxOC4wNUwyMTMuNDQgMjg5LjUxWiIgZmlsbD0iIzAwMUYzOCIvPgo8cGF0aCBkPSJNMjk3LjA2IDMxOC4xM0MzMDAuNCAzMTkuMzIgMzAyLjk5IDMyMS4yMiAzMDQuODMgMzIzLjgyQzMwNi42NyAzMjYuNDIgMzA3LjU5IDMyOS4zNyAzMDcuNTkgMzMyLjY1QzMwNy41OSAzMzguMTQgMzA1Ljc5IDM0Mi40IDMwMi4yIDM0NS40M0MyOTguNiAzNDguNDYgMjkzLjM4IDM0OS45NyAyODYuNTQgMzQ5Ljk3SDI1OC4yN1YyODkuNTJIMjg1LjYxQzI5MS45NSAyODkuNTIgMjk2Ljc5IDI5MC45NyAzMDAuMTMgMjkzLjg1QzMwMy40NyAyOTYuNzQgMzA1LjE0IDMwMC41IDMwNS4xNCAzMDUuMTRDMzA1LjE0IDMxMS4wOCAzMDIuNDUgMzE1LjQxIDI5Ny4wOCAzMTguMTNNMjg0Ljc1IDI5OC45NEgyNjkuNzJWMzE0LjczSDI4NC43NUMyODcuNjQgMzE0LjczIDI4OS44NiAzMTQuMDEgMjkxLjQyIDMxMi41NkMyOTIuOTcgMzExLjEyIDI5My43NSAzMDkuMTIgMjkzLjc1IDMwNi41OEMyOTMuNzUgMzA0LjIgMjkyLjk2IDMwMi4zMyAyOTEuMzcgMzAwLjk4QzI4OS43OCAyOTkuNjIgMjg3LjU4IDI5OC45NCAyODQuNzUgMjk4Ljk0Wk0yNjkuNzMgMzQwLjQ1SDI4NS45NUMyODkuMTggMzQwLjQ1IDI5MS42NCAzMzkuNzMgMjkzLjM0IDMzOC4yOEMyOTUuMDQgMzM2Ljg0IDI5NS44OSAzMzQuNzkgMjk1Ljg5IDMzMi4xM0MyOTUuODkgMzI2LjUzIDI5Mi41OCAzMjMuNzIgMjg1Ljk2IDMyMy43MkgyNjkuNzRWMzQwLjQ0SDI2OS43M1YzNDAuNDVaIiBmaWxsPSIjMDAxRjM4Ii8+CjxwYXRoIGQ9Ik0zNzEuMjYgMzQ5Ljk2SDM1OS4xMkwzNTQuMzcgMzM2LjYzSDMyNy44OEwzMjMuMjEgMzQ5Ljk2SDMxMS40MUwzMzQuNDIgMjg5LjUxSDM0Ny44M0wzNzEuMjYgMzQ5Ljk2Wk0zMzMuNjUgMzIwLjMzTDMzMS4zNiAzMjYuODdIMzUwLjg5TDM0OC41MSAzMjAuMzNDMzQ2LjUzIDMxNS4yOSAzNDQuMDQgMzA4LjM5IDM0MS4wNCAyOTkuNjJDMzM4LjY2IDMwNi42NCAzMzYuMiAzMTMuNTQgMzMzLjY1IDMyMC4zM1oiIGZpbGw9IiMwMDFGMzgiLz4KPHBhdGggZD0iTTQxOS4wNiAzMDUuMzFWMjg5LjUySDQzMC4xVjM0OS45N0g0MTcuMzZMNDAwLjU1IDMyMy4wNkMzOTUuOTYgMzE1Ljg3IDM5Mi4zNyAzMDkuODIgMzg5Ljc3IDMwNC44OUMzOTAgMzExLjIzIDM5MC4xMSAzMjAuOTkgMzkwLjExIDMzNC4xOFYzNDkuOTdIMzc5LjA3VjI4OS41MkgzOTEuODFMNDA4LjcgMzE2LjQzQzQxMy43NCAzMjQuNTggNDE3LjMgMzMwLjY0IDQxOS40IDMzNC42QzQxOS4xNyAzMjYuMTcgNDE5LjA2IDMxNi40IDQxOS4wNiAzMDUuMzFaIiBmaWxsPSIjMDAxRjM4Ii8+CjxwYXRoIGQ9Ik00NDMuNDIgMjg5LjUxSDQ2Ni41MUM0NzYuMTkgMjg5LjUxIDQ4My42NiAyOTIuMiA0ODguOTIgMjk3LjU4QzQ5NC4xOCAzMDMuMDEgNDk2LjgxIDExMC40IDQ5Ni44MSAzMTkuNzRDNDk2LjgxIDMyOS4wOCA0OTQuMTggMzM2LjQ0IDQ4OC45MiAzNDEuODFDNDgzLjY2IDM0Ny4yNCA0NzYuMTkgMzQ5Ljk2IDQ2Ni41MSAzNDkuOTZINDQzLjQyVjI4OS41MVpNNDU0Ljg4IDM0MC4xMkg0NjYuMzRDNDcyLjI4IDM0MC4xMiA0NzYuODcgMzM4LjM4IDQ4MC4wOSAzMzQuOUM0ODMuMzIgMzMxLjQyIDQ4NC45MyAzMjYuMzcgNDg0LjkzIDMxOS43NUM0ODQuOTMgMzEzLjEzIDQ4My4zMiAzMDguMDggNDgwLjA5IDMwNC42QzQ3Ni44NyAzMDEuMTIgNDcyLjI4IDI5OS4zOCA0NjYuMzQgMjk5LjM4SDQ1NC44OFYzNDAuMTNWMzQwLjEyWiIgZmlsbD0iIzAwMUYzOCIvPgo8cGF0aCBkPSJNNTg4LjMgMS4zMjAwMVYzODIuODgiIHN0cm9rZT0iIzAwMUYzOCIgc3Ryb2tlLXdpZHRoPSIyLjYzIiBzdHJva2UtbGluZWNhcD0icm91bmQiIHN0cm9rZS1saW5lam9pbj0icm91bmQiLz4KPGcgc3R5bGU9Im1peC1ibGVuZC1tb2RlOm11bHRpcGx5Ij4KPHBhdGggZD0iTTc1NC41NSAyOS4zN0M3MjUuNDEgMjkuMzcgNzAwLjYyIDM4LjI2IDY4MC4xNyA1Ni4wNUM2NTkuNzMgNzMuODMgNjQ5LjUgOTYuNTggNjQ5LjUgMTI0LjMySDcxMi41M0M3MTIuNTMgMTExLjQ0IDcxNS44OSAxMDAuNzkgNzIyLjYxIDkyLjM5QzcyOS4zNCA4My45OSA3MzkuOTggNzkuNzggNzU0LjU1IDc5Ljc4Qzc1OS40NyA3OS43OCA3NjMuOTIgODAuMiA3NjcuODkgODEuMDVMNzc2LjAzIDMwLjk5Qzc2OS4xOSAyOS45MSA3NjIuMDMgMjkuMzcgNzU0LjU1IDI5LjM3Wk04NDkuMSA4My4zN0g4MjYuNjZMODE1Ljg3IDE0Ny4yOEM4MjYuMjQgMTM5LjYyIDgzOS4xOSAxMzUuMDggODU0LjcyIDEzMy42N0M4NTQuODIgMTMzLjA3IDg1NC45MSAxMzIuNDcgODU0Ljk4IDEzMS44OUM4NTUuNTMgMTI3LjQxIDg1NS44MSAxMjIuNzkgODU1LjgxIDExOC4wM0M4NTUuODEgMTA1LjMzIDg1My41OCA5My43NyA4NDkuMSA4My4zOFY4My4zN1pNNzk3LjggMzI5Ljg5Qzc4MC4yMyAzMjIuNyA3NjUuNSAzMTEuODQgNzUzLjYxIDI5Ny4zMkM3NTEuODkgMjk1LjIyIDc1MC4yOSAyOTMuMDkgNzQ4LjgzIDI5MC45SDc0NC44OEw3NDcuMDggMjg4LjJDNzM5LjUzIDI3NS45OCA3MzUuNzYgMjYyLjQ5IDczNS43NiAyNDcuNzJINzgwLjE3TDc4MS40MyAyNDYuMThDNzgzLjY3IDI0My4zOCA3ODguMjEgMjM3LjkyIDc5NS4wOSAyMjkuNzlDODAxLjk1IDIyMS42NyA4MDYuNDQgMjE2LjM1IDgwOC41NCAyMTMuODJDODEwLjI5IDIxMS43MiA4MTMuMTkgMjA4LjA3IDgxNy4yMSAyMDIuOUg3NDguMDhDNzM3LjIgMjE1LjI5IDcyOS40MiAyMjQuMzkgNzI0LjcyIDIzMC4yMUw2NDcuODIgMzI1LjE2VjM0MC4yOEg4NDguOTlDODMxLjI3IDM0MC4wNyA4MTQuMiAzMzYuNiA3OTcuODEgMzI5Ljg4SDc5Ny44VjMyOS44OVoiIGZpbGw9IiM0QThCQzkiLz4KPC9nPgo8ZyBzdHlsZT0ibWl4LWJsZW5kLW1vZGU6bXVsdGlwbHkiPgo8cGF0aCBkPSJNNzg1LjQyIDkwLjUxQzc4MS4yOSA4NS44MyA3NzUuNDUgODIuNjcgNzY3Ljg5IDgxLjA3TDc0OC4wOCAyMDIuOTFDNzQ4LjIxIDIwMi43OCA3NDguMzIgMjAyLjYzIDc0OC40NSAyMDIuNDlDNzU5LjUgMTg5Ljg4IDc2Ny41OCAxODAuMzYgNzcyLjYgMTczLjkyQzc3Ny42NCAxNjcuNDcgNzgyLjYyIDE1OC45OSA3ODcuNTIgMTQ4LjVDNzkyLjQzIDEzNy45OCA3OTQuODggMTI3Ljg0IDc5NC44OCAxMTguMDNDNzk0Ljg4IDEwNi44NCA3OTEuNzMgOTcuNjQgNzg1LjQxIDkwLjUxSDc4NS40MlpNNzc2LjAzIDMwLjk5Qzc5NS43NyAzNC4xMiA4MTIuNzcgNDEuNzcgODI3LjAzIDUzLjk1QzgzNi45NiA2Mi40MiA4NDQuMzEgNzIuMjQgODQ5LjEgODMuMzhIOTQ2LjE5VjMxSDc3Ni4wM1YzMC45OVpNOTM4LjcxIDE2MS4yQzkxOS42MyAxNDIuNTIgODk1LjI3IDEzMy4xOSA4NjUuNjYgMTMzLjE5Qzg2MS44NyAxMzMuMTkgODU4LjIyIDEzMy4zNCA4NTQuNzEgMTMzLjY4Qzg1NC4wNSAxMzcuNjIgODUyLjg5IDE0MS43MiA4NTEuMTkgMTQ1Ljk3Qzg0OS4yMiAxNTAuODggODQ3LjYxIDE1NS4wNyA4NDYuMzYgMTU4LjU4Qzg0NS4wOSAxNjIuMDkgODQyLjU2IDE2Ni43MSA4MzguNzkgMTcyLjQ0QzgzNS4wMSAxNzguMTkgODMyLjIxIDE4Mi40NiA4MzAuMzkgMTg1LjI1QzgyOC41NyAxODguMDYgODI1IDE5Mi44OSA4MTkuNjggMTk5Ljc1QzgxOC44MiAyMDAuODYgODE3Ljk5IDIwMS45MSA4MTcuMjIgMjAyLjlIODE3LjIzQzgzMC41MSAxODkuNjIgODM5LjEyIDE4Mi45OCA4NTkuNiAxODIuOThDODc3LjMxIDE4Mi45OCA4ODMuNzMgMTg3LjgyIDg5Mi40NSAxOTcuNUM5MDEuMTYgMjA3LjE4IDkwNS41MiAyMTkyLjkxIDkwNS41MiAyMzUuNjlDOTA1LjUyIDI1MC4wNyA5MDEuMSAyNjIuODcgODkyLjI1IDI3NC4wOEM4ODMuNCAyODUuMjcgODcwLjI1IDI5MC44OCA4NTIuODIgMjkwLjg4QzgzOS4yNSAyOTAuODggODI3LjQzIDI4Ny4wMiA4MTcuMzMgMjc5LjI3QzgwNy4yMyAyNzEuNTIgODAyLjE4IDI2MS4wMSA4MDIuMTggMjQ3LjcySDc4MC4xOEw3NDcuMDkgMjg4LjJDNzQ3LjY1IDI4OS4xMSA3NDguMjMgMjkwLjAyIDc0OC44NCAyOTAuOUg4NTkuNkw4NTEuMiAzNDAuMjhDODUxLjc0IDM0MC4yOCA4NTIuMjcgMzQwLjI5IDg1Mi44MiAzNDAuMjhDODg2LjMgMzQwIDkxMy43NiAzMzAuMzkgOTM1LjIgMzExLjQzQzk1Ni42NCAyOTIuNDkgOTY3LjM3IDI2Ny4zNyA5NjcuMzcgMjM2LjExQzk2Ny4zNyAyMDQuODUgOTU3LjgyIDE3OS44NiA5MzguNzIgMTYxLjE5SDkzOC43MVYxNjEuMloiIGZpbGw9IiM0QThCQzkiLz4KPC9nPgo8cGF0aCBkPSJNNzQ4LjgzIDI5MC45SDc0NC44OEw3NDcuMDggMjg4LjJDNzQ3LjY0IDI4OS4xMSA3NDguMjIgMjkwLjAyIDc0OC44MyAyOTAuOVoiIGZpbGw9IiMwMDFGMzgiLz4KPHBhdGggZD0iTTg1OS41OSAyOTAuOUw4NTEuMTggMzQwLjI3SDg0OC45OEM4MzEuMjYgMzQwLjA2IDgxNC4yIDMzNi41OSA3OTcuOCAzMjkuODhDNzgwLjIzIDMyMi42OSA3NjUuNSAzMTEuODMgNzUzLjYxIDI5Ny4zQzc1MS44OSAyOTUuMiA3NTAuMjkgMjkzLjA3IDc0OC44MyAyOTAuODhIODU5LjU5VjI5MC45WiIgZmlsbD0iIzAwMUYzOCIvPgo8cGF0aCBkPSJNNzgwLjE3IDI0Ny43M0w3NDcuMDggMjg4LjJDNzM5LjUyIDI3NS45OCA3MzUuNzYgMjYyLjQ5IDczNS43NiAyNDcuNzNINzgwLjE3WiIgZmlsbD0iIzAwMUYzOCIvPgo8cGF0aCBkPSJNODU0LjcgMTMzLjY3Qzg1NC4wNSAxMzcuNjEgODUyLjg4IDE0MS43MiA4NTEuMTggMTQ1Ljk2Qzg0OS4yMiAxNTAuODcgODQ3LjYxIDE1NS4wNyA4NDYuMzUgMTU4LjU3Qzg0NS4wOCAxNjIuMDcgODQyLjU2IDE2Ni42OSA4MzguNzkgMTcyLjQzQzgzNS4wMSAxNzguMTcgODMyLjIxIDE4Mi40NSA4MzAuMzggMTg1LjI0QzgyOC41NiAxODguMDUgODI0Ljk4IDE5Mi44NyA4MTkuNjcgMTk5LjczQzgxOC44MSAyMDAuODQgODE3Ljk5IDIwMS44OSA4MTcuMjEgMjAyLjg5SDc0OC4wOUM3NDguMjIgMjAyLjc1IDc0OC4zNCAyMDIuNjEgNzQ4LjQ2IDIwMi40N0M3NTkuNTIgMTg5Ljg2IDc2Ny41OCAxODAuMzQgNzcyLjYyIDE3My45Qzc3Ny42NiAxNjcuNDUgNzgyLjY0IDE1OC45OCA3ODcuNTMgMTQ4LjQ4Qzc5Mi40NCAxMzcuOTcgNzk0Ljg5IDEyNy44MyA3OTQuODkgMTE4LjAxQzc5NC44OSAxMDYuODEgNzkxLjc0IDk3LjYzIDc4NS40MyA5MC40OUM3ODEuMyA4NS44MSA3NzUuNDYgODIuNjYgNzY3LjkgODEuMDVMNzc2LjAzIDMwLjk5Qzc5NS43NyAzNC4xMiA4MTIuNzggNDEuNzYgODI3LjAzIDUzLjk0QzgzNi45NSA2Mi40MiA4NDQuMzEgNzIuMjMgODQ5LjEgODMuMzdIODI2LjY2TDgxNS44NyAxNDcuMjhDODI2LjI0IDEzOS42MiA4MzkuMTkgMTM1LjA3IDg1NC43MiAxMzMuNjZIODU0LjdWMTMzLjY3WiIgZmlsbD0iIzAwMUYzOCIvPgo8cGF0aCBkPSJNMTA1NC45NCAzMS43SDEwNzYuOThMMTA1MS4xNCA4MC43MVYxMDcuMzhIMTAzMC4xNFY4MC43MUwxMDA0LjMgMzEuN0gxMDI2LjMzTDEwNDAuNjQgNjIuMzhMMTA1NC45NSAzMS43SDEwNTQuOTRaIiBmaWxsPSIjNEE4QkM5Ii8+CjxwYXRoIGQ9Ik0xMTAwLjg2IDMxLjdWMTA3LjM4SDEwNzkuODZWMzEuN0gxMTAwLjg2WiIgZmlsbD0iIzRBOEJDOSIvPgo8cGF0aCBkPSJNMTE2OC40MSAxMDcuMzlIMTExMS4wNlY0OC42SDExMDUuOTFMMTExMS4wNiAzMS43MUgxMTMyLjA2VjkwLjVIMTE2OC40MVYxMDcuMzlaIiBmaWxsPSIjNEE4QkM5Ii8+CjxwYXRoIGQ9Ik0xMjM3LjUgNjkuNDlDMTIzNy41IDg4LjQ0IDEyMjUuOTcgMTA3LjM4IDEyMDIuOCAxMDcuMzhIMTE3NC4wN1Y0OC41OUgxMTY4LjkyTDExNzQuMDcgMzEuN0gxMjAyLjhDMTIyNS45NyAzMS43IDEyMzcuNSA1MC41NCAxMjM3LjUgNjkuNDlaTTExOTYuNjMgOTAuNUMxMjA4Ljg4IDkwLjUgMTIxNC45NiA4MC4xIDEyMTQuOTYgNjkuN0MxMjE0Ljk2IDU5LjMgMTIwOC43OCA0OC41OSAxMTk2LjYzIDQ4LjU5SDExOTUuMDlWOTAuNUgxMTk2LjYzWiIgZmlsbD0iIzRBOEJDOSIvPgo8cGF0aCBkPSJNMTI2NC4xNyAzMS43VjEwNy4zOEgxMjQzLjE2VjMxLjdIMTI2NC4xN1oiIGZpbGw9IiM0QThCQzkiLz4KPHBhdGggZD0iTTEzMzUuNzMgNTcuNTVDMTMzNS43MyA2Ny4wMiAxMzMwLjk5IDc2LjI5IDEzMjEuMTEgODBMMTMzOS4zNCAxMDcuMzlIMTMxNS4yNUwxMzAwLjAxIDgxLjY1SDEyOTUuMzhWMTA3LjM5SDEyNzQuMzhWNDguNkgxMjY5LjIzTDEyNzQuMzggMzEuNzFIMTMxMS4zNEMxMzI3LjIgMzEuNzEgMTMzNS43NCA0NC43OSAxMzM1Ljc0IDU3LjU1SDEzMzUuNzNNMTI5NS4zNyA2NC43NkgxMzA1LjE1QzEzMTAuNSA2NC43NiAxMzEzLjE4IDYwLjc0IDEzMTMuMTggNTYuNzNDMTMxMy4xOCA1Mi43MiAxMzEwLjUgNDguNiAxMzA1LjE1IDQ4LjZIMTI5NS4zN1Y2NC43N1Y2NC43NloiIGZpbGw9IiM0QThCQzkiLz4KPC9zdmc+Cg==" style="width: 191px; height: 55px; margin-bottom: 10px;">
                        <h2 style="margin: 0; font-size: 22px; color: #1F76AC; font-weight: 800; font-family: 'Inter', sans-serif;">Açık Sipariş Raporu</h2>
                    </td>
                    <td style="vertical-align: top; text-align: right; font-family: 'Inter', sans-serif; width: 60%;">
                        <p style="margin: 0 0 5px 0; font-size: 14px; color: #718096;">Rapor Tarihi: <strong>${today}</strong></p>
                        <p style="margin: 0 0 10px 0; font-size: 14px; color: #718096;">Toplam Açık Bakiye: <strong style="color: #1F76AC; font-size: 18px;">${totalVal}</strong></p>
                        
                        <table style="width: auto; float: right; margin-top: 5px; background: #EDF2F7; border-collapse: collapse; border-radius: 6px;">
                            <tr>
                                <td style="padding: 8px 12px; text-align: left;">
                                    <div style="font-size: 11px; color: #4A5568; font-weight: 600; margin-bottom: 3px; text-transform: uppercase;">Baz Alınan Kurlar</div>
                                    <div style="font-size: 12px; color: #2D3748;">
                                        <strong>1 EUR =</strong> ${FORMATTER.number(r.USD)} USD <span style="color:#CBD5E0">|</span> ${FORMATTER.number(r.GBP)} GBP <span style="color:#CBD5E0">|</span> ${FORMATTER.number(r.TRY)} TL
                                    </div>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        `;

        const dashboard = document.getElementById('mainDashboard');
        dashboard.insertBefore(printHeader, dashboard.firstChild);

        // 3. html2canvas ile ekran görüntüsü alma
        const fullWidth = dashboard.scrollWidth;
        const fullHeight = dashboard.scrollHeight;
        const canvas = await html2canvas(dashboard, {
            scale: 2,
            useCORS: true,
            logging: false,
            backgroundColor: '#F5F7FA',
            windowWidth: fullWidth,
            width: fullWidth,
            height: fullHeight,
            scrollX: 0,
            scrollY: 0
        });

        // 4. Temizlik (Gizlenen tabloyu ve filtreleri geri getir, başlığı sil)
        dashboard.removeChild(printHeader);
        if (filterBar) filterBar.style.display = 'flex';
        if (tableControls) tableControls.style.display = 'flex';
        if (tableSection) tableSection.style.display = 'block'; // Tabloyu geri getir
        rankingLists.forEach(el => { el.style.maxHeight = ''; el.style.overflow = ''; });

        // 5. PDF Oluşturma
        const imgData = canvas.toDataURL('image/jpeg', 0.95);
        const { jsPDF } = window.jspdf;

        // Convert canvas pixels → mm (96 px/inch, 25.4 mm/inch → 1px = 25.4/96 mm)
        // canvas.width already includes scale:2, so divide by scale first
        const PX_TO_MM = 25.4 / 96;
        const pdfWidth  = (canvas.width  / 2) * PX_TO_MM;
        const pdfHeight = (canvas.height / 2) * PX_TO_MM;

        const pdf = new jsPDF({ orientation: pdfWidth > pdfHeight ? 'l' : 'p', unit: 'mm', format: [pdfWidth, pdfHeight] });
        pdf.addImage(imgData, 'JPEG', 0, 0, pdfWidth, pdfHeight);
        pdf.save(`Ayik_Band_Acik_Siparisler_${today}.pdf`);

    } catch (err) {
        console.error("PDF Export Hatası:", err);
        alert("PDF oluşturulurken bir hata oluştu. Lütfen tekrar deneyin.");
    } finally {
        btn.innerHTML = originalText;
        btn.disabled = false;
    }
}

