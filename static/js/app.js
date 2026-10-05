// ==========================================================================
// ALIFIL FINANCIAL INTELLIGENCE - APP LOGIC & VISUALIZATION ENGINE
// ==========================================================================

let state = {
  data: [],            // List data hasil analisis
  filteredData: [],    // Data setelah filter & pencarian
  selectedFilter: 'all', // 'all', 'bigfour', 'rotation', 'healthy', 'loss'
  searchQuery: '',
  sortCol: 'roa_percent',
  sortAsc: false,
  queuedFiles: [],
  roaChartInstance: null,
  kapChartInstance: null
};

// Inisialisasi Aplikasi Saat Halaman Selesai Dimuat
document.addEventListener('DOMContentLoaded', () => {
  initDropzone();
  initTableEvents();
  initFilterButtons();
  initExportButtons();
  
  // Muat Demo secara otomatis agar dashboard langsung hidup dan memukau!
  loadDemoData();
});

// ==========================================================================
// 1. DRAG & DROP DAN UPLOAD HANDLERS
// ==========================================================================
function initDropzone() {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('file-input');

  if (!dropzone || !fileInput) return;

  dropzone.addEventListener('click', () => fileInput.click());

  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('dragover');
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files.length > 0) {
      handleFilesAdded(e.dataTransfer.files);
    }
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleFilesAdded(e.target.files);
    }
  });

  const btnProcessQueue = document.getElementById('btn-process-queue');
  if (btnProcessQueue) {
    btnProcessQueue.addEventListener('click', processUploadQueue);
  }
}

function handleFilesAdded(fileList) {
  for (let i = 0; i < fileList.length; i++) {
    const file = fileList[i];
    if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
      // Cek apakah sudah ada di queue
      if (!state.queuedFiles.some(f => f.name === file.name && f.size === file.size)) {
        state.queuedFiles.push(file);
      }
    } else {
      alert(`File "${file.name}" dilewati karena bukan file Excel (.xlsx atau .xls).`);
    }
  }
  renderFileQueue();
}

function renderFileQueue() {
  const queueBox = document.getElementById('file-queue-box');
  const queueContainer = document.getElementById('queue-container');
  const queueCount = document.getElementById('queue-count');

  if (!queueBox || !queueContainer) return;

  if (state.queuedFiles.length === 0) {
    queueContainer.style.display = 'none';
    queueBox.innerHTML = '';
    return;
  }

  queueContainer.style.display = 'block';
  queueCount.textContent = state.queuedFiles.length;

  queueBox.innerHTML = state.queuedFiles.map((file, idx) => `
    <div class="queue-item">
      <div class="queue-file-info">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="#818CF8" stroke-width="2">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
          <polyline points="14 2 14 8 20 8"></polyline>
          <line x1="16" y1="13" x2="8" y2="13"></line>
          <line x1="16" y1="17" x2="8" y2="17"></line>
          <polyline points="10 9 9 9 8 9"></polyline>
        </svg>
        <span style="font-weight: 500;">${file.name}</span>
        <span style="color: var(--text-dim); font-size: 11px;">(${(file.size / 1024).toFixed(1)} KB)</span>
      </div>
      <button class="queue-remove-btn" onclick="removeQueueItem(${idx})" title="Hapus file">
        ✕
      </button>
    </div>
  `).join('');
}

window.removeQueueItem = function(idx) {
  state.queuedFiles.splice(idx, 1);
  renderFileQueue();
};

async function processUploadQueue() {
  if (state.queuedFiles.length === 0) return;

  const btn = document.getElementById('btn-process-queue');
  const originalText = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = `<span class="spinner"></span> Memproses ${state.queuedFiles.length} File...`;

  const formData = new FormData();
  state.queuedFiles.forEach(file => {
    formData.append('files', file);
  });

  try {
    const res = await fetch('/api/upload', {
      method: 'POST',
      body: formData
    });
    const result = await res.json();

    if (result.success) {
      // Gabungkan dengan data yang sudah ada (update atau tambahkan)
      result.data.forEach(newItem => {
        const existingIdx = state.data.findIndex(d => d.ticker === newItem.ticker && d.period === newItem.period);
        if (existingIdx >= 0) {
          state.data[existingIdx] = newItem;
        } else {
          state.data.push(newItem);
        }
      });

      state.queuedFiles = [];
      renderFileQueue();
      applyFilterAndRender();
      showToast(`Sukses mengekstrak ${result.processed_count} laporan keuangan!`);
    } else {
      alert("Gagal memproses file: " + (result.detail || "Terjadi kesalahan"));
    }
  } catch (err) {
    console.error("Upload error:", err);
    alert("Terjadi kesalahan koneksi saat mengunggah file.");
  } finally {
    btn.disabled = false;
    btn.innerHTML = originalText;
  }
}

// ==========================================================================
// 2. DEMO DATA LOADER & RESET
// ==========================================================================
async function loadDemoData() {
  const btn = document.getElementById('btn-load-demo');
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = `<span>⏳</span> Memuat Data IDX...`;
  }

  try {
    const res = await fetch('/api/demo-data');
    const result = await res.json();
    if (result.success) {
      state.data = result.data;
      applyFilterAndRender();
      showToast("6 Data Laporan Keuangan BEI Realistis Berhasil Dimuat!");
    }
  } catch (err) {
    console.error("Error loading demo data:", err);
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = `
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <polygon points="5 3 19 12 5 21 5 3"></polygon>
        </svg>
        Muat Demo IDX (6 Emiten)
      `;
    }
  }
}

function resetAllData() {
  if (confirm("Apakah Anda yakin ingin mengosongkan semua data hasil analisis?")) {
    state.data = [];
    state.queuedFiles = [];
    renderFileQueue();
    applyFilterAndRender();
    showToast("Data analisis telah dibersihkan.");
  }
}

// ==========================================================================
// 3. FILTER, PENCARIAN, DAN SORTING
// ==========================================================================
function initFilterButtons() {
  const filterBtns = document.querySelectorAll('.filter-btn');
  filterBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      filterBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      state.selectedFilter = btn.dataset.filter;
      applyFilterAndRender();
    });
  });

  const searchInput = document.getElementById('search-input');
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      state.searchQuery = e.target.value.toLowerCase().trim();
      applyFilterAndRender();
    });
  }
}

function applyFilterAndRender() {
  let list = [...state.data];

  // 1. Terapkan Filter Kategori
  if (state.selectedFilter === 'bigfour') {
    list = list.filter(item => item.is_big_four);
  } else if (state.selectedFilter === 'rotation') {
    list = list.filter(item => item.auditor_switching && item.auditor_switching.includes('Berganti'));
  } else if (state.selectedFilter === 'healthy') {
    list = list.filter(item => item.roa_percent !== null && item.roa_percent >= 5.0);
  } else if (state.selectedFilter === 'loss') {
    list = list.filter(item => item.roa_percent !== null && item.roa_percent < 0);
  }

  // 2. Terapkan Pencarian Teks
  if (state.searchQuery) {
    list = list.filter(item => {
      const ticker = (item.ticker || '').toLowerCase();
      const entity = (item.entity_name || '').toLowerCase();
      const kap = (item.current_auditor || '').toLowerCase();
      const partner = (item.audit_partner || '').toLowerCase();
      const priorKap = (item.prior_auditor || '').toLowerCase();
      return ticker.includes(state.searchQuery) ||
             entity.includes(state.searchQuery) ||
             kap.includes(state.searchQuery) ||
             partner.includes(state.searchQuery) ||
             priorKap.includes(state.searchQuery);
    });
  }

  // 3. Terapkan Pengurutan (Sorting)
  list.sort((a, b) => {
    let valA = a[state.sortCol];
    let valB = b[state.sortCol];

    if (valA === null || valA === undefined) valA = -999999999;
    if (valB === null || valB === undefined) valB = -999999999;

    if (typeof valA === 'string') {
      return state.sortAsc ? valA.localeCompare(valB) : valB.localeCompare(valA);
    }
    return state.sortAsc ? valA - valB : valB - valA;
  });

  state.filteredData = list;

  updateKPIs();
  renderTable();
  updateCharts();
}

function initTableEvents() {
  const headers = document.querySelectorAll('.data-table th[data-sort]');
  headers.forEach(th => {
    th.addEventListener('click', () => {
      const col = th.dataset.sort;
      if (state.sortCol === col) {
        state.sortAsc = !state.sortAsc;
      } else {
        state.sortCol = col;
        state.sortAsc = false;
      }
      applyFilterAndRender();
    });
  });
}

// ==========================================================================
// 4. RENDER KPI CARDS
// ==========================================================================
function updateKPIs() {
  const totalEmiten = state.data.length;
  
  let validRoaCount = 0;
  let totalRoaSum = 0;
  let totalAssetsSum = 0;
  let bigFourCount = 0;
  let rotationCount = 0;

  state.data.forEach(item => {
    if (item.roa_percent !== null) {
      totalRoaSum += item.roa_percent;
      validRoaCount++;
    }
    if (item.total_assets_scaled) {
      totalAssetsSum += item.total_assets_scaled;
    }
    if (item.is_big_four) {
      bigFourCount++;
    }
    if (item.auditor_switching && item.auditor_switching.includes('Berganti')) {
      rotationCount++;
    }
  });

  const avgRoa = validRoaCount > 0 ? (totalRoaSum / validRoaCount).toFixed(2) + '%' : '0.00%';
  const bigFourPct = totalEmiten > 0 ? Math.round((bigFourCount / totalEmiten) * 100) + '%' : '0%';
  
  // Format Total Aset dalam Triliun Rupiah
  const assetsTriliun = (totalAssetsSum / 1000000000000).toFixed(1);

  document.getElementById('kpi-total-emiten').textContent = totalEmiten;
  document.getElementById('kpi-avg-roa').textContent = avgRoa;
  document.getElementById('kpi-total-assets').textContent = `Rp ${Number(assetsTriliun).toLocaleString('id-ID')} T`;
  document.getElementById('kpi-bigfour-share').textContent = bigFourPct;
  document.getElementById('kpi-auditor-switch').textContent = rotationCount;
}

// ==========================================================================
// 5. RENDER TABEL UTAMA
// ==========================================================================
function renderTable() {
  const tbody = document.getElementById('table-body');
  const emptyState = document.getElementById('empty-state');
  const countBadge = document.getElementById('result-count');

  if (!tbody) return;

  if (countBadge) {
    countBadge.textContent = `${state.filteredData.length} dari ${state.data.length} Emiten`;
  }

  if (state.filteredData.length === 0) {
    tbody.innerHTML = '';
    if (emptyState) emptyState.style.display = 'block';
    return;
  }

  if (emptyState) emptyState.style.display = 'none';

  tbody.innerHTML = state.filteredData.map((item, idx) => {
    const isSwitched = item.auditor_switching && item.auditor_switching.includes('Berganti');
    const switchBadgeClass = isSwitched ? 'pill-switched' : 'pill-same';
    const switchIcon = isSwitched ? '⚡' : '✓';

    const kapTierClass = item.is_big_four ? 'pill-bigfour' : 'pill-nonbigfour';

    // Format nilai aset dan laba dalam Rupiah riil yang ringkas
    const formattedAssetReal = item.total_assets_scaled !== null 
      ? formatRupiahCompact(item.total_assets_scaled) 
      : 'N/A';

    const formattedIncomeReal = item.net_income_scaled !== null 
      ? formatRupiahCompact(item.net_income_scaled) 
      : 'N/A';

    const roaColorClass = `roa-${item.roa_badge_color || 'blue'}`;

    return `
      <tr>
        <td style="color: var(--text-dim); text-align: center;">${idx + 1}</td>
        <td>
          <span class="ticker-badge">${item.ticker}</span>
          <div class="company-title">${escapeHtml(item.entity_name)}</div>
        </td>
        <td style="font-size: 12px; color: var(--text-muted);">${escapeHtml(item.period)}</td>
        <td>
          <div style="font-weight: 600; color: var(--text-main);">${escapeHtml(item.current_auditor)}</div>
          <span class="pill-badge ${kapTierClass}" style="margin-top: 4px;">
            ${item.kap_tier}
          </span>
        </td>
        <td>
          <div style="font-weight: 500;">${escapeHtml(item.audit_partner)}</div>
          <div style="font-size: 11px; color: var(--text-dim);">Signing Partner</div>
        </td>
        <td>
          <div style="font-size: 12px; color: var(--text-muted);">${escapeHtml(item.prior_auditor)}</div>
          <span class="pill-badge ${switchBadgeClass}" style="margin-top: 4px;">
            ${switchIcon} ${item.auditor_switching}
          </span>
        </td>
        <td style="text-align: right;">
          <div style="font-family: var(--font-mono); font-weight: 600; color: #FFFFFF;">
            ${formattedAssetReal}
          </div>
          <div style="font-size: 10px; color: var(--text-dim);">
            ${item.rounding_text}
          </div>
        </td>
        <td style="text-align: right;">
          <div style="font-family: var(--font-mono); font-weight: 600; color: ${item.net_income_scaled < 0 ? '#FB7185' : '#34D399'};">
            ${formattedIncomeReal}
          </div>
          <div style="font-size: 10px; color: var(--text-dim);">
            ${item.net_income_scaled < 0 ? 'Rugi Bersih' : 'Laba Bersih'}
          </div>
        </td>
        <td style="text-align: right;">
          <div class="roa-cell ${roaColorClass}">
            ${item.roa_formatted}
          </div>
          <span class="pill-badge" style="background: rgba(255,255,255,0.06); font-size: 10px; padding: 2px 6px;">
            ${item.roa_status.split('(')[0].trim()}
          </span>
        </td>
        <td style="text-align: center;">
          <button class="btn btn-outline btn-sm" onclick="openDetailModal('${item.ticker}')" title="Lihat Jejak Audit & Perhitungan">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <circle cx="11" cy="11" r="8"></circle>
              <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
              <line x1="11" y1="8" x2="11" y2="14"></line>
              <line x1="8" y1="11" x2="14" y2="11"></line>
            </svg>
            Detail
          </button>
        </td>
      </tr>
    `;
  }).join('');
}

// ==========================================================================
// 6. CHART.JS VISUALISASI
// ==========================================================================
function updateCharts() {
  if (typeof Chart === 'undefined') return;

  const validItems = state.data.filter(d => d.roa_percent !== null);

  // 1. Chart ROA Comparison (Horizontal Bar)
  const ctxRoa = document.getElementById('chart-roa');
  if (ctxRoa) {
    const labels = validItems.map(d => d.ticker);
    const roaValues = validItems.map(d => d.roa_percent);
    const bgColors = validItems.map(d => {
      if (d.roa_percent >= 10) return 'rgba(16, 185, 129, 0.85)';
      if (d.roa_percent >= 5) return 'rgba(6, 182, 212, 0.85)';
      if (d.roa_percent >= 2) return 'rgba(59, 130, 246, 0.85)';
      if (d.roa_percent >= 0) return 'rgba(245, 158, 11, 0.85)';
      return 'rgba(244, 63, 94, 0.85)';
    });

    if (state.roaChartInstance) {
      state.roaChartInstance.destroy();
    }

    state.roaChartInstance = new Chart(ctxRoa, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [{
          label: 'ROA (%)',
          data: roaValues,
          backgroundColor: bgColors,
          borderRadius: 8,
          borderSkipped: false
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (context) => {
                const item = validItems[context.dataIndex];
                return `${item.entity_name}: ${context.raw.toFixed(2)}% (${item.roa_status})`;
              }
            }
          }
        },
        scales: {
          x: {
            grid: { color: 'rgba(255, 255, 255, 0.05)' },
            ticks: { color: '#94A3B8', font: { family: 'Plus Jakarta Sans', weight: 'bold' } }
          },
          y: {
            grid: { color: 'rgba(255, 255, 255, 0.05)' },
            ticks: {
              color: '#94A3B8',
              callback: (val) => `${val}%`
            }
          }
        }
      }
    });
  }

  // 2. Chart KAP Market Share (Doughnut)
  const ctxKap = document.getElementById('chart-kap');
  if (ctxKap) {
    const kapCounts = {};
    state.data.forEach(d => {
      let aud = d.current_auditor || 'Lainnya';
      // Singkat nama KAP untuk visualisasi rapi
      if (aud.includes('Pricewaterhouse') || aud.includes('Tanudiredja')) aud = 'PwC (Tanudiredja)';
      else if (aud.includes('Ernst & Young') || aud.includes('Purwantono')) aud = 'EY (Purwantono)';
      else if (aud.includes('Deloitte') || aud.includes('Imelda')) aud = 'Deloitte (Imelda)';
      else if (aud.includes('KPMG') || aud.includes('Siddharta')) aud = 'KPMG (Siddharta)';
      else aud = 'KAP Non-Big Four';
      
      kapCounts[aud] = (kapCounts[aud] || 0) + 1;
    });

    if (state.kapChartInstance) {
      state.kapChartInstance.destroy();
    }

    state.kapChartInstance = new Chart(ctxKap, {
      type: 'doughnut',
      data: {
        labels: Object.keys(kapCounts),
        datasets: [{
          data: Object.values(kapCounts),
          backgroundColor: [
            '#4F46E5', '#06B6D4', '#10B981', '#F59E0B', '#8B5CF6', '#64748B'
          ],
          borderColor: '#0D121F',
          borderWidth: 3
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            labels: {
              color: '#94A3B8',
              font: { size: 11, family: 'Plus Jakarta Sans' },
              padding: 12
            }
          }
        },
        cutout: '65%'
      }
    });
  }
}

// ==========================================================================
// 7. MODAL DETAIL & AUDIT TRAIL
// ==========================================================================
window.openDetailModal = function(ticker) {
  const item = state.data.find(d => d.ticker === ticker);
  if (!item) return;

  const modal = document.getElementById('detail-modal');
  const body = document.getElementById('modal-body-content');
  const title = document.getElementById('modal-title');

  title.innerHTML = `
    <span style="color: #60A5FA; font-family: var(--font-mono);">${item.ticker}</span> - 
    <span>${escapeHtml(item.entity_name)}</span>
  `;

  const trail = item.audit_trail || {};

  body.innerHTML = `
    <!-- Kartu Perhitungan Matematika ROA -->
    <div class="formula-card">
      <div class="formula-title">📐 Rumus & Perhitungan Return on Assets (ROA)</div>
      <div style="font-size: 13px; color: var(--text-muted); margin-bottom: 8px;">
        $$\\text{ROA (\\%)} = \\left( \\frac{\\text{Laba Bersih}}{\\text{Total Aset}} \\right) \\times 100\\%$$
      </div>
      <div class="formula-math">${item.roa_formula}</div>
      <div style="margin-top: 8px; font-size: 12px; color: #6EE7B7;">
        Status Kesehatan Finansial: <strong>${item.roa_status}</strong>
      </div>
    </div>

    <!-- Ringkasan Angka Finansial -->
    <div class="audit-grid">
      <div class="audit-item">
        <div class="audit-label">Total Aset (Sesuai Laporan)</div>
        <div class="audit-val">${item.total_assets_raw !== null ? item.total_assets_raw.toLocaleString('id-ID') : 'N/A'}</div>
        <div style="font-size: 11px; color: var(--text-dim); margin-top: 4px;">Skala: ${item.rounding_text}</div>
      </div>
      <div class="audit-item">
        <div class="audit-label">Total Aset Riil (Penuh Rp)</div>
        <div class="audit-val" style="color: #60A5FA;">Rp ${item.total_assets_scaled !== null ? item.total_assets_scaled.toLocaleString('id-ID') : 'N/A'}</div>
        <div style="font-size: 11px; color: var(--text-dim); margin-top: 4px;">(${formatRupiahCompact(item.total_assets_scaled)})</div>
      </div>
      <div class="audit-item">
        <div class="audit-label">Laba Bersih (Sesuai Laporan)</div>
        <div class="audit-val">${item.net_income_raw !== null ? item.net_income_raw.toLocaleString('id-ID') : 'N/A'}</div>
        <div style="font-size: 11px; color: var(--text-dim); margin-top: 4px;">Periode Berjalan</div>
      </div>
      <div class="audit-item">
        <div class="audit-label">Laba Bersih Riil (Penuh Rp)</div>
        <div class="audit-val" style="color: ${item.net_income_scaled < 0 ? '#FB7185' : '#34D399'};">
          Rp ${item.net_income_scaled !== null ? item.net_income_scaled.toLocaleString('id-ID') : 'N/A'}
        </div>
        <div style="font-size: 11px; color: var(--text-dim); margin-top: 4px;">(${formatRupiahCompact(item.net_income_scaled)})</div>
      </div>
    </div>

    <!-- Informasi Audit & Tata Kelola KAP -->
    <div style="background: rgba(255, 255, 255, 0.02); border: 1px solid var(--border-subtle); border-radius: 14px; padding: 18px;">
      <h4 style="font-size: 13px; font-weight: 700; color: #A5B4FC; margin-bottom: 12px; text-transform: uppercase;">
        🛡️ Tata Kelola Audit & Auditor Independen
      </h4>
      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 14px; font-size: 12px;">
        <div>
          <span style="color: var(--text-dim);">Auditor Tahun Berjalan:</span>
          <div style="font-weight: 600; color: var(--text-main); margin-top: 2px;">${escapeHtml(item.current_auditor)}</div>
          <span class="pill-badge ${item.is_big_four ? 'pill-bigfour' : 'pill-nonbigfour'}" style="margin-top: 4px;">${item.kap_tier}</span>
        </div>
        <div>
          <span style="color: var(--text-dim);">Signing Partner (Rekan Audit):</span>
          <div style="font-weight: 600; color: #FFFFFF; margin-top: 2px;">${escapeHtml(item.audit_partner)}</div>
        </div>
        <div>
          <span style="color: var(--text-dim);">Auditor Tahun Sebelumnya:</span>
          <div style="font-weight: 600; color: var(--text-muted); margin-top: 2px;">${escapeHtml(item.prior_auditor)}</div>
        </div>
        <div>
          <span style="color: var(--text-dim);">Status Rotasi KAP:</span>
          <div style="margin-top: 2px;">
            <span class="pill-badge ${item.auditor_switching && item.auditor_switching.includes('Berganti') ? 'pill-switched' : 'pill-same'}">
              ${item.auditor_switching}
            </span>
          </div>
        </div>
      </div>
    </div>

    <!-- Jejak Audit Ekstraksi e-XBRL (Traceability) -->
    <div style="background: rgba(15, 23, 42, 0.5); border: 1px solid var(--border-subtle); border-radius: 14px; padding: 18px;">
      <h4 style="font-size: 13px; font-weight: 700; color: #94A3B8; margin-bottom: 12px; text-transform: uppercase;">
        🔍 Jejak Pencocokan Baris & Sheet e-XBRL
      </h4>
      <div style="display: flex; flex-direction: column; gap: 10px; font-size: 12px; font-family: var(--font-mono);">
        <div style="display: flex; justify-content: space-between; border-bottom: 1px dashed var(--border-subtle); padding-bottom: 6px;">
          <span style="color: var(--text-dim);">Sheet Info Umum:</span>
          <span style="color: #60A5FA;">${trail.general_sheet || 'N/A'}</span>
        </div>
        <div style="display: flex; justify-content: space-between; border-bottom: 1px dashed var(--border-subtle); padding-bottom: 6px;">
          <span style="color: var(--text-dim);">Sheet Posisi Keuangan (Neraca):</span>
          <span style="color: #60A5FA;">${trail.balance_sheet || 'N/A'} (Baris ${trail.balance_row_matched || 'N/A'})</span>
        </div>
        <div style="display: flex; justify-content: space-between; border-bottom: 1px dashed var(--border-subtle); padding-bottom: 6px;">
          <span style="color: var(--text-dim);">Label Baris Total Aset:</span>
          <span style="color: var(--text-main); max-width: 380px; text-align: right;">${escapeHtml(trail.balance_label_matched || 'N/A')}</span>
        </div>
        <div style="display: flex; justify-content: space-between; border-bottom: 1px dashed var(--border-subtle); padding-bottom: 6px;">
          <span style="color: var(--text-dim);">Sheet Laba Rugi:</span>
          <span style="color: #60A5FA;">${trail.income_sheet || 'N/A'} (Baris ${trail.income_row_matched || 'N/A'})</span>
        </div>
        <div style="display: flex; justify-content: space-between; padding-bottom: 4px;">
          <span style="color: var(--text-dim);">Label Baris Laba Bersih:</span>
          <span style="color: var(--text-main); max-width: 380px; text-align: right;">${escapeHtml(trail.income_label_matched || 'N/A')}</span>
        </div>
      </div>
    </div>
  `;

  modal.classList.add('active');
};

window.closeDetailModal = function() {
  const modal = document.getElementById('detail-modal');
  if (modal) modal.classList.remove('active');
};

// ==========================================================================
// 8. EXPORT LAPORAN (EXCEL & CSV)
// ==========================================================================
function initExportButtons() {
  const btnExcel = document.getElementById('btn-export-excel');
  const btnCsv = document.getElementById('btn-export-csv');

  if (btnExcel) {
    btnExcel.addEventListener('click', async () => {
      if (state.data.length === 0) {
        alert("Tidak ada data untuk diekspor.");
        return;
      }
      try {
        btnExcel.disabled = true;
        const res = await fetch('/api/export/excel', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ data: state.data })
        });
        if (res.ok) {
          const blob = await res.blob();
          downloadBlob(blob, 'Analisis_Laporan_Keuangan_ROA_IDX.xlsx');
          showToast("File Excel berhasil diunduh!");
        } else {
          alert("Gagal mengekspor file Excel.");
        }
      } catch (err) {
        console.error(err);
      } finally {
        btnExcel.disabled = false;
      }
    });
  }

  if (btnCsv) {
    btnCsv.addEventListener('click', async () => {
      if (state.data.length === 0) {
        alert("Tidak ada data untuk diekspor.");
        return;
      }
      try {
        btnCsv.disabled = true;
        const res = await fetch('/api/export/csv', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ data: state.data })
        });
        if (res.ok) {
          const blob = await res.blob();
          downloadBlob(blob, 'Analisis_Laporan_Keuangan_ROA_IDX.csv');
          showToast("File CSV berhasil diunduh!");
        } else {
          alert("Gagal mengekspor file CSV.");
        }
      } catch (err) {
        console.error(err);
      } finally {
        btnCsv.disabled = false;
      }
    });
  }
}

function downloadBlob(blob, filename) {
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.style.display = 'none';
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  window.URL.revokeObjectURL(url);
}

// Download Sample File Tunggal
window.downloadSample = function(ticker) {
  window.location.href = `/api/download-sample/${ticker}`;
};

// ==========================================================================
// 9. HELPER UTILITIES
// ==========================================================================
function formatRupiahCompact(num) {
  if (num === null || num === undefined) return 'N/A';
  const abs = Math.abs(num);
  const sign = num < 0 ? '-' : '';

  if (abs >= 1000000000000) {
    return `${sign}Rp ${(abs / 1000000000000).toLocaleString('id-ID', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} Triliun`;
  } else if (abs >= 1000000000) {
    return `${sign}Rp ${(abs / 1000000000).toLocaleString('id-ID', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} Miliar`;
  } else if (abs >= 1000000) {
    return `${sign}Rp ${(abs / 1000000).toLocaleString('id-ID', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} Juta`;
  }
  return `${sign}Rp ${abs.toLocaleString('id-ID')}`;
}

function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function showToast(msg) {
  let toast = document.getElementById('app-toast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'app-toast';
    toast.style.position = 'fixed';
    toast.style.bottom = '24px';
    toast.style.right = '24px';
    toast.style.background = '#1E293B';
    toast.style.color = '#FFFFFF';
    toast.style.border = '1px solid rgba(16, 185, 129, 0.5)';
    toast.style.padding = '12px 20px';
    toast.style.borderRadius = '12px';
    toast.style.boxShadow = '0 10px 30px rgba(0, 0, 0, 0.5)';
    toast.style.fontSize = '13px';
    toast.style.fontWeight = '600';
    toast.style.zIndex = '9999';
    toast.style.transition = 'all 0.3s ease';
    toast.style.display = 'flex';
    toast.style.alignItems = 'center';
    toast.style.gap = '8px';
    document.body.appendChild(toast);
  }

  toast.innerHTML = `<span>✓</span> ${msg}`;
  toast.style.opacity = '1';
  toast.style.transform = 'translateY(0)';

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
  }, 3500);
}
