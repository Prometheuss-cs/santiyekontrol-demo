/**
 * app.js
 * ŞantiyeKontrol arayüz kontrolcüsü.
 */
(function () {
  'use strict';

  const { parseSpreadsheetFile, createXlsxBlob, triggerDownload } = window.XlsxLite;
  const { guessColumnMapping, reconcileDocuments, calculateHakedis } = window.ReconcileEngine;

  // --- Durum (State) ---
  const state = {
    irsaliyeFile: null, // { fileName, headers, records }
    faturaFile: null,
    irsMap: {},
    fatMap: {},
    reconciled: null, // { rows, summary, analytics }
    activeFilter: 'ALL',
    activeSite: 'ALL',
    searchQuery: '',
    hakedisItems: [],
  };

  // --- Yardımcılar ---
  const $ = (id) => document.getElementById(id);

  function esc(str) {
    return String(str === null || str === undefined ? '' : str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function fmtTL(val) {
    return (Number(val) || 0).toLocaleString('tr-TR', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }) + ' ₺';
  }

  function fmtNum(val) {
    return (Number(val) || 0).toLocaleString('tr-TR', {
      minimumFractionDigits: 0,
      maximumFractionDigits: 2,
    });
  }

  let toastTimer = null;
  function showToast(msg) {
    const el = $('toast');
    el.textContent = msg;
    el.classList.remove('hidden');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => el.classList.add('hidden'), 3400);
  }

  // --- Örnek İnşaat Muhasebesi Verisi ---
  function getSampleDatasets() {
    const irsaliyeler = [
      { 'İrsaliye No': 'IRS-2026-101', 'Tarih': '02.11.2026', 'Şantiye': 'Güneşli Konutları', 'Tedarikçi': 'Anadolu Hazır Beton A.Ş.', 'Malzeme Cinsi': 'C30 Hazır Beton', 'Birim': 'm³', 'Miktar': 40, 'Birim Fiyat': 2850, 'Toplam Tutar': 114000 },
      { 'İrsaliye No': 'IRS-2026-102', 'Tarih': '03.11.2026', 'Şantiye': 'Güneşli Konutları', 'Tedarikçi': 'Anadolu Hazır Beton A.Ş.', 'Malzeme Cinsi': 'C30 Hazır Beton', 'Birim': 'm³', 'Miktar': 24, 'Birim Fiyat': 2850, 'Toplam Tutar': 68400 },
      { 'İrsaliye No': 'IRS-2026-103', 'Tarih': '04.11.2026', 'Şantiye': 'Kartal Rezidans', 'Tedarikçi': 'DemirÇelik Yapı San. Tic.', 'Malzeme Cinsi': 'Ø14 Nervürlü İnşaat Demiri', 'Birim': 'Ton', 'Miktar': 15, 'Birim Fiyat': 24200, 'Toplam Tutar': 363000 },
      { 'İrsaliye No': 'IRS-2026-104', 'Tarih': '05.11.2026', 'Şantiye': 'Güneşli Konutları', 'Tedarikçi': 'Anadolu Hazır Beton A.Ş.', 'Malzeme Cinsi': 'C30 Hazır Beton', 'Birim': 'm³', 'Miktar': 30, 'Birim Fiyat': 2850, 'Toplam Tutar': 85500 },
      { 'İrsaliye No': 'IRS-2026-105', 'Tarih': '06.11.2026', 'Şantiye': 'Ataşehir İş Merkezi', 'Tedarikçi': 'Kaya Agrega & Yapı Malz.', 'Malzeme Cinsi': 'Kırma Taş Mıcır (15-25 mm)', 'Birim': 'Ton', 'Miktar': 85, 'Birim Fiyat': 420, 'Toplam Tutar': 35700 },
      { 'İrsaliye No': 'IRS-2026-106', 'Tarih': '07.11.2026', 'Şantiye': 'Kartal Rezidans', 'Tedarikçi': 'Marmara Nakliyat Ltd.', 'Malzeme Cinsi': 'Hafriyat & Moloz Nakliyesi', 'Birim': 'Sefer', 'Miktar': 12, 'Birim Fiyat': 3200, 'Toplam Tutar': 38400 },
      { 'İrsaliye No': 'IRS-2026-107', 'Tarih': '08.11.2026', 'Şantiye': 'Ataşehir İş Merkezi', 'Tedarikçi': 'Anadolu Hazır Beton A.Ş.', 'Malzeme Cinsi': 'C35 Hazır Beton + Pompa', 'Birim': 'm³', 'Miktar': 52, 'Birim Fiyat': 3150, 'Toplam Tutar': 163800 },
      { 'İrsaliye No': 'IRS-2026-108', 'Tarih': '09.11.2026', 'Şantiye': 'Kartal Rezidans', 'Tedarikçi': 'DemirÇelik Yapı San. Tic.', 'Malzeme Cinsi': 'Ø16 Nervürlü İnşaat Demiri', 'Birim': 'Ton', 'Miktar': 18, 'Birim Fiyat': 24200, 'Toplam Tutar': 435600 },
      { 'İrsaliye No': 'IRS-2026-109', 'Tarih': '10.11.2026', 'Şantiye': 'Güneşli Konutları', 'Tedarikçi': 'Kaya Agrega & Yapı Malz.', 'Malzeme Cinsi': 'Yıkanmış Sıva Kumu', 'Birim': 'Ton', 'Miktar': 60, 'Birim Fiyat': 480, 'Toplam Tutar': 28800 },
      { 'İrsaliye No': 'IRS-2026-110', 'Tarih': '11.11.2026', 'Şantiye': 'Ataşehir İş Merkezi', 'Tedarikçi': 'DemirÇelik Yapı San. Tic.', 'Malzeme Cinsi': 'Q188 Çelik Hasır', 'Birim': 'Adet', 'Miktar': 140, 'Birim Fiyat': 690, 'Toplam Tutar': 96600 },
      { 'İrsaliye No': 'IRS-2026-111', 'Tarih': '11.11.2026', 'Şantiye': 'Güneşli Konutları', 'Tedarikçi': 'Marmara Nakliyat Ltd.', 'Malzeme Cinsi': 'Malzeme Vinç & Nakliye', 'Birim': 'Saat', 'Miktar': 16, 'Birim Fiyat': 1950, 'Toplam Tutar': 31200 },
      { 'İrsaliye No': 'IRS-2026-112', 'Tarih': '12.11.2026', 'Şantiye': 'Kartal Rezidans', 'Tedarikçi': 'Anadolu Hazır Beton A.Ş.', 'Malzeme Cinsi': 'C30 Hazır Beton', 'Birim': 'm³', 'Miktar': 45, 'Birim Fiyat': 2850, 'Toplam Tutar': 128250 },
      { 'İrsaliye No': 'IRS-2026-113', 'Tarih': '13.11.2026', 'Şantiye': 'Ataşehir İş Merkezi', 'Tedarikçi': 'Kaya Agrega & Yapı Malz.', 'Malzeme Cinsi': 'Bims Blok 20cm', 'Birim': 'Adet', 'Miktar': 2500, 'Birim Fiyat': 22, 'Toplam Tutar': 55000 },
      { 'İrsaliye No': 'IRS-2026-114', 'Tarih': '14.11.2026', 'Şantiye': 'Ataşehir İş Merkezi', 'Tedarikçi': 'Kaya Agrega & Yapı Malz.', 'Malzeme Cinsi': 'CEM I 42.5 Çimento', 'Birim': 'Torba', 'Miktar': 160, 'Birim Fiyat': 175, 'Toplam Tutar': 28000 },
      { 'İrsaliye No': 'IRS-2026-115', 'Tarih': '14.11.2026', 'Şantiye': 'Güneşli Konutları', 'Tedarikçi': 'DemirÇelik Yapı San. Tic.', 'Malzeme Cinsi': 'Ø12 Nervürlü İnşaat Demiri', 'Birim': 'Ton', 'Miktar': 6, 'Birim Fiyat': 24200, 'Toplam Tutar': 145200 },
    ];

    const faturalar = [
      { 'Fatura No': 'FAT-4410', 'İrsaliye No': 'IRS-2026-101', 'Tarih': '05.11.2026', 'Şantiye': 'Güneşli Konutları', 'Tedarikçi': 'Anadolu Hazır Beton A.Ş.', 'Malzeme Cinsi': 'C30 Hazır Beton', 'Birim': 'm³', 'Miktar': 40, 'Birim Fiyat': 2850, 'Toplam Tutar': 114000 },
      { 'Fatura No': 'FAT-4412', 'İrsaliye No': 'IRS-2026-102', 'Tarih': '05.11.2026', 'Şantiye': 'Güneşli Konutları', 'Tedarikçi': 'Anadolu Hazır Beton A.Ş.', 'Malzeme Cinsi': 'C30 Hazır Beton', 'Birim': 'm³', 'Miktar': 24, 'Birim Fiyat': 2850, 'Toplam Tutar': 68400 },
      { 'Fatura No': 'FAT-7021', 'İrsaliye No': 'IRS-2026-103', 'Tarih': '06.11.2026', 'Şantiye': 'Kartal Rezidans', 'Tedarikçi': 'DemirÇelik Yapı San. Tic.', 'Malzeme Cinsi': 'Ø14 Nervürlü İnşaat Demiri', 'Birim': 'Ton', 'Miktar': 15, 'Birim Fiyat': 24200, 'Toplam Tutar': 363000 },
      // HATA 1: Miktar farkı (30 m³ yerine 36 m³ fatura kesilmiş)
      { 'Fatura No': 'FAT-4418', 'İrsaliye No': 'IRS-2026-104', 'Tarih': '07.11.2026', 'Şantiye': 'Güneşli Konutları', 'Tedarikçi': 'Anadolu Hazır Beton A.Ş.', 'Malzeme Cinsi': 'C30 Hazır Beton', 'Birim': 'm³', 'Miktar': 36, 'Birim Fiyat': 2850, 'Toplam Tutar': 102600 },
      // HATA 2: Mükerrer fatura (IRS-2026-102 ikinci kez fatura edilmiş)
      { 'Fatura No': 'FAT-4419', 'İrsaliye No': 'IRS-2026-102', 'Tarih': '08.11.2026', 'Şantiye': 'Güneşli Konutları', 'Tedarikçi': 'Anadolu Hazır Beton A.Ş.', 'Malzeme Cinsi': 'C30 Hazır Beton', 'Birim': 'm³', 'Miktar': 24, 'Birim Fiyat': 2850, 'Toplam Tutar': 68400 },
      { 'Fatura No': 'FAT-3190', 'İrsaliye No': 'IRS-2026-105', 'Tarih': '08.11.2026', 'Şantiye': 'Ataşehir İş Merkezi', 'Tedarikçi': 'Kaya Agrega & Yapı Malz.', 'Malzeme Cinsi': 'Kırma Taş Mıcır (15-25 mm)', 'Birim': 'Ton', 'Miktar': 85, 'Birim Fiyat': 420, 'Toplam Tutar': 35700 },
      { 'Fatura No': 'FAT-8830', 'İrsaliye No': 'IRS-2026-106', 'Tarih': '09.11.2026', 'Şantiye': 'Kartal Rezidans', 'Tedarikçi': 'Marmara Nakliyat Ltd.', 'Malzeme Cinsi': 'Hafriyat & Moloz Nakliyesi', 'Birim': 'Sefer', 'Miktar': 12, 'Birim Fiyat': 3200, 'Toplam Tutar': 38400 },
      { 'Fatura No': 'FAT-4425', 'İrsaliye No': 'IRS-2026-107', 'Tarih': '10.11.2026', 'Şantiye': 'Ataşehir İş Merkezi', 'Tedarikçi': 'Anadolu Hazır Beton A.Ş.', 'Malzeme Cinsi': 'C35 Hazır Beton + Pompa', 'Birim': 'm³', 'Miktar': 52, 'Birim Fiyat': 3150, 'Toplam Tutar': 163800 },
      // HATA 3: Sözleşme birim fiyatı 24.200 iken 25.400 TL kesilmiş
      { 'Fatura No': 'FAT-7033', 'İrsaliye No': 'IRS-2026-108', 'Tarih': '11.11.2026', 'Şantiye': 'Kartal Rezidans', 'Tedarikçi': 'DemirÇelik Yapı San. Tic.', 'Malzeme Cinsi': 'Ø16 Nervürlü İnşaat Demiri', 'Birim': 'Ton', 'Miktar': 18, 'Birim Fiyat': 25400, 'Toplam Tutar': 457200 },
      { 'Fatura No': 'FAT-3204', 'İrsaliye No': 'IRS-2026-109', 'Tarih': '11.11.2026', 'Şantiye': 'Güneşli Konutları', 'Tedarikçi': 'Kaya Agrega & Yapı Malz.', 'Malzeme Cinsi': 'Yıkanmış Sıva Kumu', 'Birim': 'Ton', 'Miktar': 60, 'Birim Fiyat': 480, 'Toplam Tutar': 28800 },
      { 'Fatura No': 'FAT-7040', 'İrsaliye No': 'IRS-2026-110', 'Tarih': '12.11.2026', 'Şantiye': 'Ataşehir İş Merkezi', 'Tedarikçi': 'DemirÇelik Yapı San. Tic.', 'Malzeme Cinsi': 'Q188 Çelik Hasır', 'Birim': 'Adet', 'Miktar': 140, 'Birim Fiyat': 690, 'Toplam Tutar': 96600 },
      { 'Fatura No': 'FAT-8832', 'İrsaliye No': 'IRS-2026-111', 'Tarih': '12.11.2026', 'Şantiye': 'Güneşli Konutları', 'Tedarikçi': 'Marmara Nakliyat Ltd.', 'Malzeme Cinsi': 'Malzeme Vinç & Nakliye', 'Birim': 'Saat', 'Miktar': 16, 'Birim Fiyat': 1950, 'Toplam Tutar': 31200 },
      { 'Fatura No': 'FAT-4439', 'İrsaliye No': 'IRS-2026-112', 'Tarih': '13.11.2026', 'Şantiye': 'Kartal Rezidans', 'Tedarikçi': 'Anadolu Hazır Beton A.Ş.', 'Malzeme Cinsi': 'C30 Hazır Beton', 'Birim': 'm³', 'Miktar': 45, 'Birim Fiyat': 2850, 'Toplam Tutar': 128250 },
      { 'Fatura No': 'FAT-3211', 'İrsaliye No': 'IRS-2026-113', 'Tarih': '14.11.2026', 'Şantiye': 'Ataşehir İş Merkezi', 'Tedarikçi': 'Kaya Agrega & Yapı Malz.', 'Malzeme Cinsi': 'Bims Blok 20cm', 'Birim': 'Adet', 'Miktar': 2500, 'Birim Fiyat': 22, 'Toplam Tutar': 55000 },
      // HATA 4: İrsaliyesi şantiyede olmayan fatura
      { 'Fatura No': 'FAT-8834', 'İrsaliye No': 'IRS-2026-999', 'Tarih': '14.11.2026', 'Şantiye': 'Kartal Rezidans', 'Tedarikçi': 'Marmara Nakliyat Ltd.', 'Malzeme Cinsi': 'Ekstra Kamyon Seferi (Onaysız)', 'Birim': 'Sefer', 'Miktar': 5, 'Birim Fiyat': 3700, 'Toplam Tutar': 18500 },
    ];

    return { irsaliyeler, faturalar };
  }

  // --- Sekme (Tab) Yönetimi ---
  function initTabs() {
    const buttons = document.querySelectorAll('.tab-btn');
    const panels = document.querySelectorAll('.tab-panel');
    buttons.forEach((btn) => {
      btn.addEventListener('click', () => {
        const targetId = btn.dataset.tab;
        buttons.forEach((b) => b.classList.toggle('active', b === btn));
        panels.forEach((p) => p.classList.toggle('active', p.id === targetId));
      });
    });
  }

  // --- Dosya Yükleme & Sütun Eşleştirme ---
  function setupDropzone(dropId, inputId, pickBtnId, kind) {
    const drop = $(dropId);
    const input = $(inputId);
    const pickBtn = $(pickBtnId);

    pickBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      input.click();
    });

    drop.addEventListener('click', () => input.click());

    drop.addEventListener('dragover', (e) => {
      e.preventDefault();
      drop.classList.add('dragover');
    });

    drop.addEventListener('dragleave', () => drop.classList.remove('dragover'));

    drop.addEventListener('drop', (e) => {
      e.preventDefault();
      drop.classList.remove('dragover');
      if (e.dataTransfer.files && e.dataTransfer.files[0]) {
        handleFileSelected(e.dataTransfer.files[0], kind);
      }
    });

    input.addEventListener('change', () => {
      if (input.files && input.files[0]) {
        handleFileSelected(input.files[0], kind);
        input.value = '';
      }
    });
  }

  async function handleFileSelected(file, kind) {
    try {
      const parsed = await parseSpreadsheetFile(file);
      if (kind === 'irsaliye') {
        state.irsaliyeFile = parsed;
        state.irsMap = guessColumnMapping(parsed.headers, 'irsaliye');
        $('dropIrsaliye').classList.add('loaded');
        $('statusIrsaliye').textContent = `${parsed.fileName} — ${parsed.records.length} satır (${parsed.sheetName})`;
      } else {
        state.faturaFile = parsed;
        state.fatMap = guessColumnMapping(parsed.headers, 'fatura');
        $('dropFatura').classList.add('loaded');
        $('statusFatura').textContent = `${parsed.fileName} — ${parsed.records.length} satır (${parsed.sheetName})`;
      }

      renderMappingSelectors();
      if (state.irsaliyeFile && state.faturaFile) {
        runReconciliation();
        showToast('İki dosya başarıyla karşılaştırıldı.');
      } else {
        showToast(`${parsed.fileName} yüklendi. Şimdi diğer dosyayı seçin.`);
      }
    } catch (err) {
      alert(`Dosya okunurken hata oluştu:\n${err.message}`);
    }
  }

  const FIELD_LABELS = {
    irsaliyeNo: 'İrsaliye No *',
    faturaNo: 'Fatura No',
    tarih: 'Tarih',
    santiye: 'Şantiye / Proje',
    tedarikci: 'Tedarikçi / Cari',
    malzeme: 'Malzeme / Açıklama',
    birim: 'Birim',
    miktar: 'Miktar',
    birimFiyat: 'Birim Fiyat',
    tutar: 'Toplam Tutar',
  };

  function renderMappingGrid(containerId, headers, currentMap, onChange) {
    const el = $(containerId);
    if (!headers || headers.length === 0) {
      el.innerHTML = '<div class="cell-sub">Önce dosya yükleyin.</div>';
      return;
    }

    el.innerHTML = Object.keys(currentMap)
      .map((field) => {
        const opts = ['<option value="">(Yok / Boş)</option>']
          .concat(
            headers.map((h) => `<option value="${esc(h)}" ${currentMap[field] === h ? 'selected' : ''}>${esc(h)}</option>`)
          )
          .join('');
        return `
          <label class="map-field">
            <span>${FIELD_LABELS[field] || field}</span>
            <select class="input-control" data-field="${field}">${opts}</select>
          </label>
        `;
      })
      .join('');

    el.querySelectorAll('select').forEach((sel) => {
      sel.addEventListener('change', () => {
        onChange(sel.dataset.field, sel.value);
      });
    });
  }

  function renderMappingSelectors() {
    renderMappingGrid('mapIrsaliyeGrid', state.irsaliyeFile?.headers, state.irsMap, (f, v) => {
      state.irsMap[f] = v;
    });
    renderMappingGrid('mapFaturaGrid', state.faturaFile?.headers, state.fatMap, (f, v) => {
      state.fatMap[f] = v;
    });
  }

  function loadDemoData() {
    const { irsaliyeler, faturalar } = getSampleDatasets();
    const irsHeaders = Object.keys(irsaliyeler[0]);
    const fatHeaders = Object.keys(faturalar[0]);

    state.irsaliyeFile = {
      fileName: 'ornek_santiye_irsaliyeleri.xlsx',
      sheetName: 'İrsaliyeler',
      headers: irsHeaders,
      records: irsaliyeler,
    };
    state.faturaFile = {
      fileName: 'ornek_tedarikci_faturalari.xlsx',
      sheetName: 'Faturalar',
      headers: fatHeaders,
      records: faturalar,
    };

    state.irsMap = guessColumnMapping(irsHeaders, 'irsaliye');
    state.fatMap = guessColumnMapping(fatHeaders, 'fatura');

    $('dropIrsaliye').classList.add('loaded');
    $('statusIrsaliye').textContent = `Örnek İrsaliye Verisi — ${irsaliyeler.length} satır`;
    $('dropFatura').classList.add('loaded');
    $('statusFatura').textContent = `Örnek Fatura Verisi — ${faturalar.length} satır`;

    renderMappingSelectors();
    runReconciliation();
    showToast('Örnek inşaat verisi yüklendi ve 4 kritik hata yakalandı!');
  }

  // --- Çapraz Kontrol ve Tablo Render ---
  function runReconciliation() {
    if (!state.irsaliyeFile || !state.faturaFile) {
      showToast('Lütfen önce hem irsaliye hem fatura dosyasını yükleyin.');
      return;
    }

    state.reconciled = reconcileDocuments(
      state.irsaliyeFile.records,
      state.faturaFile.records,
      state.irsMap,
      state.fatMap
    );

    state.activeFilter = 'ALL';
    state.activeSite = 'ALL';
    state.searchQuery = '';
    $('searchInput').value = '';

    $('reconcilePlaceholder').classList.add('hidden');
    $('reconcileResults').classList.remove('hidden');

    renderReconcileKpis();
    populateSiteFilter();
    updateFilterPills();
    renderReconcileTable();
    renderAnalytics();
    updatePrintMeta();
  }

  function renderReconcileKpis() {
    const s = state.reconciled.summary;
    $('kpiTotalRows').textContent = s.toplamKalem;
    $('kpiMatchedSub').textContent = `${s.eslesenSayisi} kalem birebir eşleşti`;
    $('kpiErrorRows').textContent = s.hataliSayisi;
    const errRate = s.toplamKalem > 0 ? Math.round((s.hataliSayisi / s.toplamKalem) * 100) : 0;
    $('kpiErrorSub').textContent = `Kontrol edilenlerin %${errRate}'i sorunlu`;
    $('kpiRiskAmount').textContent = fmtTL(s.fazlaFaturalananRisk);
    $('kpiPendingRows').textContent = s.bekleyenSayisi;
    $('kpiPendingSub').textContent = `${fmtTL(s.bekleyenIrsaliyeTutar)} tutarında açık irsaliye`;
  }

  function populateSiteFilter() {
    const sites = state.reconciled.analytics.santiyeler.map((s) => s.name);
    const select = $('siteFilter');
    select.innerHTML = '<option value="ALL">Tüm Şantiyeler</option>' +
      sites.map((name) => `<option value="${esc(name)}">${esc(name)}</option>`).join('');
  }

  function updateFilterPills() {
    const rows = state.reconciled.rows;
    const countOf = (key) => {
      if (key === 'ALL') return rows.length;
      if (key === 'ERRORS') return rows.filter((r) => r.durum !== 'EŞLEŞTİ' && r.durum !== 'FATURA BEKLİYOR').length;
      if (key === 'MİKTAR') return rows.filter((r) => r.durum.includes('MİKTAR')).length;
      if (key === 'FİYAT') return rows.filter((r) => r.durum.includes('FİYAT')).length;
      return rows.filter((r) => r.durum === key).length;
    };

    const labels = {
      ALL: 'Tümü',
      ERRORS: 'Sadece Sorunlular',
      MİKTAR: 'Miktar Farkı',
      FİYAT: 'Fiyat Farkı',
      'İRSALİYE YOK': 'İrsaliye Yok',
      'MÜKERRER FATURA': 'Mükerrer',
      'FATURA BEKLİYOR': 'Fatura Bekliyor',
      EŞLEŞTİ: 'Eşleşti',
    };

    document.querySelectorAll('#filterPills .pill').forEach((pill) => {
      const key = pill.dataset.filter;
      pill.textContent = `${labels[key] || key} (${countOf(key)})`;
      pill.classList.toggle('active', key === state.activeFilter);
    });
  }

  function getFilteredRows() {
    if (!state.reconciled) return [];
    const q = state.searchQuery.trim().toLocaleLowerCase('tr-TR');

    return state.reconciled.rows.filter((r) => {
      if (state.activeSite !== 'ALL' && r.santiye !== state.activeSite) return false;
      if (state.activeFilter === 'ERRORS') {
        if (r.durum === 'EŞLEŞTİ' || r.durum === 'FATURA BEKLİYOR') return false;
      } else if (state.activeFilter === 'MİKTAR') {
        if (!r.durum.includes('MİKTAR')) return false;
      } else if (state.activeFilter === 'FİYAT') {
        if (!r.durum.includes('FİYAT')) return false;
      } else if (state.activeFilter !== 'ALL' && r.durum !== state.activeFilter) {
        return false;
      }

      if (q) {
        const hay = `${r.irsaliyeNo} ${r.faturaNo} ${r.santiye} ${r.tedarikci} ${r.malzeme} ${r.durum}`.toLocaleLowerCase('tr-TR');
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }

  function statusBadgeClass(severity) {
    if (severity === 'ok') return 'status-ok';
    if (severity === 'danger') return 'status-danger';
    if (severity === 'warn') return 'status-warn';
    return 'status-info';
  }

  function renderReconcileTable() {
    const rows = getFilteredRows();
    const tbody = $('reconcileTableBody');
    const empty = $('emptyTableState');

    if (rows.length === 0) {
      tbody.innerHTML = '';
      empty.classList.remove('hidden');
      return;
    }

    empty.classList.add('hidden');
    tbody.innerHTML = rows
      .map((r) => {
        const rowClass = r.severity === 'danger' ? 'row-danger' : r.severity === 'warn' ? 'row-warn' : '';
        const diffClass = r.farkTutar > 0.5 ? 'diff-pos' : r.farkTutar < -0.5 ? 'diff-neg' : 'diff-zero';
        const diffText =
          r.durum === 'FATURA BEKLİYOR'
            ? '—'
            : `${r.farkTutar > 0 ? '+' : ''}${fmtTL(r.farkTutar)}`;

        const irsSub =
          r.irsMiktar > 0 ? `${fmtNum(r.irsMiktar)} ${esc(r.birim)} × ${fmtTL(r.irsBirimFiyat)}` : 'Kayıt yok';
        const fatSub =
          r.fatMiktar > 0
            ? `${fmtNum(r.fatMiktar)} ${esc(r.birim)} × ${fmtTL(r.fatBirimFiyat)}`
            : r.durum === 'FATURA BEKLİYOR'
            ? 'Fatura gelmedi'
            : '—';

        return `
          <tr class="${rowClass}">
            <td><span class="status-badge ${statusBadgeClass(r.severity)}">${esc(r.durum)}</span></td>
            <td>
              <div class="cell-main">${esc(r.irsaliyeNo)}</div>
              <div class="cell-sub">Fat: ${esc(r.faturaNo)} ${r.tarih ? '• ' + esc(r.tarih) : ''}</div>
            </td>
            <td>
              <div class="cell-main">${esc(r.santiye)}</div>
              <div class="cell-sub">${esc(r.tedarikci)}</div>
            </td>
            <td><div class="cell-main">${esc(r.malzeme)}</div></td>
            <td class="num">
              <div class="cell-main">${r.irsTutar > 0 ? fmtTL(r.irsTutar) : '—'}</div>
              <div class="cell-sub">${irsSub}</div>
            </td>
            <td class="num">
              <div class="cell-main">${r.fatTutar > 0 ? fmtTL(r.fatTutar) : '—'}</div>
              <div class="cell-sub">${fatSub}</div>
            </td>
            <td class="num ${diffClass}">${diffText}</td>
            <td>${esc(r.aciklama)}</td>
          </tr>
        `;
      })
      .join('');
  }

  // --- Sekme 2: Şantiye & Tedarikçi Özeti ---
  function renderAnalytics() {
    if (!state.reconciled) return;
    $('analyticsPlaceholder').classList.add('hidden');
    $('analyticsContent').classList.remove('hidden');
    $('btnAnalyticsDemo').classList.add('hidden');

    const { santiyeler, kategoriler, tedarikciler } = state.reconciled.analytics;

    const maxSite = Math.max(1, ...santiyeler.map((s) => s.tutar));
    $('siteBarsContainer').innerHTML = santiyeler
      .map((s) => {
        const pct = Math.round((s.tutar / maxSite) * 100);
        return `
          <div>
            <div class="bar-item-head">
              <span class="bar-item-name">${esc(s.name)}</span>
              <span class="bar-item-val">${fmtTL(s.tutar)}</span>
            </div>
            <div class="bar-track"><div class="bar-fill" style="width:${pct}%"></div></div>
            <div class="bar-item-meta">${s.kalem} kalem • ${s.hata} hatalı fatura • Yakalanan risk: ${fmtTL(s.risk)}</div>
          </div>
        `;
      })
      .join('');

    const maxCat = Math.max(1, ...kategoriler.map((c) => c.tutar));
    $('categoryBarsContainer').innerHTML = kategoriler
      .map((c) => {
        const pct = Math.round((c.tutar / maxCat) * 100);
        return `
          <div>
            <div class="bar-item-head">
              <span class="bar-item-name">${esc(c.name)}</span>
              <span class="bar-item-val">${fmtTL(c.tutar)}</span>
            </div>
            <div class="bar-track"><div class="bar-fill amber" style="width:${pct}%"></div></div>
            <div class="bar-item-meta">${c.kalem} kalem</div>
          </div>
        `;
      })
      .join('');

    $('supplierTableBody').innerHTML = tedarikciler
      .map((t) => `
        <tr class="${t.risk > 0 ? 'row-danger' : ''}">
          <td><div class="cell-main">${esc(t.name)}</div></td>
          <td class="num">${t.kalem}</td>
          <td class="num">${fmtTL(t.irsaliyeTutar)}</td>
          <td class="num">${fmtTL(t.faturaTutar)}</td>
          <td class="num">${t.hata > 0 ? `<span class="status-badge status-danger">${t.hata} hata</span>` : '<span class="status-badge status-ok">Temiz</span>'}</td>
          <td class="num ${t.risk > 0 ? 'diff-pos' : 'diff-zero'}">${fmtTL(t.risk)}</td>
        </tr>
      `)
      .join('');
  }

  // --- Sekme 3: Taşeron Hakediş Hesaplayıcı ---
  const DEFAULT_HAKEDIS_ITEMS = [
    { ad: 'Kaba Yapı Kalıp İşçiliği (B Blok 3-4. Kat)', miktar: 1450, birim: 'm²', fiyat: 340 },
    { ad: 'İnşaat Demiri İşçiliği (Bağlama & Montaj)', miktar: 42, birim: 'Ton', fiyat: 5200 },
    { ad: 'Beton Döküm ve Vibratör İşçiliği', miktar: 310, birim: 'm³', fiyat: 185 },
  ];

  function initHakedis() {
    state.hakedisItems = DEFAULT_HAKEDIS_ITEMS.map((item) => ({ ...item }));
    renderHakedisItems();
    recalcHakedis();

    const inputs = ['hkFirma', 'hkSantiye', 'hkDonem', 'hkKdvRate', 'hkTevkifat', 'hkTeminatRate', 'hkStopajRate', 'hkAvans', 'hkDigerKesinti'];
    inputs.forEach((id) => {
      $(id).addEventListener('input', recalcHakedis);
      $(id).addEventListener('change', recalcHakedis);
    });

    $('btnAddHakedisRow').addEventListener('click', () => {
      state.hakedisItems.push({ ad: 'Yeni İş Kalemi', miktar: 1, birim: 'Adet', fiyat: 0 });
      renderHakedisItems();
      recalcHakedis();
    });

    $('btnResetHakedis').addEventListener('click', () => {
      $('hkFirma').value = 'Yıldız Kalıp & Demir İnşaat Ltd.';
      $('hkSantiye').value = 'Güneşli Konutları - B Blok';
      $('hkDonem').value = 'Hakediş #04 — Kasım 2026';
      $('hkKdvRate').value = '20';
      $('hkTevkifat').value = '0.4';
      $('hkTeminatRate').value = '5';
      $('hkStopajRate').value = '0';
      $('hkAvans').value = '35000';
      $('hkDigerKesinti').value = '4500';
      state.hakedisItems = DEFAULT_HAKEDIS_ITEMS.map((item) => ({ ...item }));
      renderHakedisItems();
      recalcHakedis();
      showToast('Hakediş örnek değerlere sıfırlandı.');
    });
  }

  function renderHakedisItems() {
    const tbody = $('hkItemsBody');
    tbody.innerHTML = state.hakedisItems
      .map((item, idx) => {
        const lineTotal = (Number(item.miktar) || 0) * (Number(item.fiyat) || 0);
        return `
          <tr>
            <td><input class="input-control hk-edit" data-idx="${idx}" data-key="ad" type="text" value="${esc(item.ad)}" /></td>
            <td><input class="input-control hk-edit num" data-idx="${idx}" data-key="miktar" type="number" step="0.01" min="0" value="${item.miktar}" /></td>
            <td><input class="input-control hk-edit" data-idx="${idx}" data-key="birim" type="text" value="${esc(item.birim)}" /></td>
            <td><input class="input-control hk-edit num" data-idx="${idx}" data-key="fiyat" type="number" step="0.01" min="0" value="${item.fiyat}" /></td>
            <td class="num cell-main" id="hkLineTotal-${idx}">${fmtTL(lineTotal)}</td>
            <td class="no-print"><button class="btn-icon-danger hk-del" data-idx="${idx}" type="button" title="Kalemi sil">×</button></td>
          </tr>
        `;
      })
      .join('');

    tbody.querySelectorAll('.hk-edit').forEach((inp) => {
      inp.addEventListener('input', () => {
        const idx = Number(inp.dataset.idx);
        const key = inp.dataset.key;
        state.hakedisItems[idx][key] = key === 'miktar' || key === 'fiyat' ? Number(inp.value) || 0 : inp.value;
        const lineTotal = state.hakedisItems[idx].miktar * state.hakedisItems[idx].fiyat;
        const cell = $(`hkLineTotal-${idx}`);
        if (cell) cell.textContent = fmtTL(lineTotal);
        recalcHakedis();
      });
    });

    tbody.querySelectorAll('.hk-del').forEach((btn) => {
      btn.addEventListener('click', () => {
        if (state.hakedisItems.length <= 1) {
          showToast('En az 1 iş kalemi bulunmalıdır.');
          return;
        }
        state.hakedisItems.splice(Number(btn.dataset.idx), 1);
        renderHakedisItems();
        recalcHakedis();
      });
    });
  }

  function getHakedisResult() {
    const brutTutar = state.hakedisItems.reduce(
      (sum, item) => sum + (Number(item.miktar) || 0) * (Number(item.fiyat) || 0),
      0
    );
    return calculateHakedis({
      brutTutar,
      kdvOrani: $('hkKdvRate').value,
      tevkifatPay: $('hkTevkifat').value,
      teminatOrani: $('hkTeminatRate').value,
      stopajOrani: $('hkStopajRate').value,
      avansMahsubu: $('hkAvans').value,
      digerKesinti: $('hkDigerKesinti').value,
    });
  }

  function recalcHakedis() {
    const res = getHakedisResult();
    $('hkSummaryHeader').textContent = `${$('hkFirma').value} — ${$('hkDonem').value}`;
    $('sumBrut').textContent = fmtTL(res.brutTutar);
    $('sumToplamKdv').textContent = fmtTL(res.toplamKdv);
    $('sumTevkifatKdv').textContent = fmtTL(res.tevkifEdilenKdv);
    $('sumTaseronKdv').textContent = fmtTL(res.taseronaKalanKdv);
    $('sumFaturaToplam').textContent = fmtTL(res.faturaOdenecekToplam);

    $('lblTeminatPct').textContent = res.teminatOrani;
    $('sumTeminat').textContent = '-' + fmtTL(res.teminatKesintisi);
    $('lblStopajPct').textContent = res.stopajOrani;
    $('sumStopaj').textContent = '-' + fmtTL(res.stopajKesintisi);
    $('sumAvans').textContent = '-' + fmtTL(res.avansMahsubu);
    $('sumDiger').textContent = '-' + fmtTL(res.digerKesinti);
    $('sumToplamKesinti').textContent = '-' + fmtTL(res.toplamKesinti);
    $('sumNetOdenecek').textContent = fmtTL(res.netOdenecek);
  }

  // --- Excel (.xlsx) Çıktı İşlemleri ---
  async function exportReconcileExcel() {
    if (!state.reconciled) return;
    const { rows, summary, analytics } = state.reconciled;

    const headerStyle = 1;
    const rowStyleFor = (sev, isNum) => {
      if (sev === 'danger') return isNum ? 4 : 3;
      if (sev === 'warn') return isNum ? 6 : 5;
      if (sev === 'ok') return isNum ? 8 : 7;
      return isNum ? 2 : 0;
    };

    const summaryRows = [
      [{ v: 'ŞantiyeKontrol — İrsaliye & Fatura Kontrol Özeti', s: 9 }, { v: '', s: 9 }],
      ['Rapor Tarihi', new Date().toLocaleDateString('tr-TR')],
      ['Toplam Kontrol Edilen Kalem', { v: summary.toplamKalem, s: 0 }],
      ['Birebir Eşleşen Kalem', { v: summary.eslesenSayisi, s: 7 }],
      ['Tespit Edilen Hatalı / Farklı Kalem', { v: summary.hataliSayisi, s: 3 }],
      ['Fatura Bekleyen İrsaliye', { v: summary.bekleyenSayisi, s: 0 }],
      ['Toplam İrsaliye Tutarı (TL)', { v: summary.toplamIrsaliyeTutar, s: 2 }],
      ['Toplam Fatura Tutarı (TL)', { v: summary.toplamFaturaTutar, s: 2 }],
      [{ v: 'Yakalanan Fazla Faturalandırma Riski (TL)', s: 9 }, { v: summary.fazlaFaturalananRisk, s: 4 }],
      [],
      [{ v: 'Sadece Sorunlu / İnceleme Gerektiren Kalemler', s: 9 }],
      ['Durum', 'İrsaliye No', 'Fatura No', 'Şantiye', 'Tedarikçi', 'Malzeme', 'İrs. Tutar', 'Fat. Tutar', 'Fark (TL)', 'Açıklama'].map((h) => ({ v: h, s: headerStyle })),
    ];

    rows
      .filter((r) => r.durum !== 'EŞLEŞTİ')
      .forEach((r) => {
        summaryRows.push([
          { v: r.durum, s: rowStyleFor(r.severity, false) },
          { v: r.irsaliyeNo, s: rowStyleFor(r.severity, false) },
          { v: r.faturaNo, s: rowStyleFor(r.severity, false) },
          { v: r.santiye, s: rowStyleFor(r.severity, false) },
          { v: r.tedarikci, s: rowStyleFor(r.severity, false) },
          { v: r.malzeme, s: rowStyleFor(r.severity, false) },
          { v: r.irsTutar, s: rowStyleFor(r.severity, true) },
          { v: r.fatTutar, s: rowStyleFor(r.severity, true) },
          { v: r.farkTutar, s: rowStyleFor(r.severity, true) },
          { v: r.aciklama, s: rowStyleFor(r.severity, false) },
        ]);
      });

    const allRows = [
      ['Durum', 'İrsaliye No', 'Fatura No', 'Tarih', 'Şantiye', 'Tedarikçi', 'Malzeme', 'Birim', 'İrs. Miktar', 'Fat. Miktar', 'İrs. B.Fiyat', 'Fat. B.Fiyat', 'İrs. Tutar', 'Fat. Tutar', 'Fark (TL)', 'Açıklama'].map((h) => ({ v: h, s: headerStyle })),
      ...rows.map((r) => [
        { v: r.durum, s: rowStyleFor(r.severity, false) },
        { v: r.irsaliyeNo, s: 0 },
        { v: r.faturaNo, s: 0 },
        { v: r.tarih, s: 0 },
        { v: r.santiye, s: 0 },
        { v: r.tedarikci, s: 0 },
        { v: r.malzeme, s: 0 },
        { v: r.birim, s: 0 },
        { v: r.irsMiktar, s: 2 },
        { v: r.fatMiktar, s: 2 },
        { v: r.irsBirimFiyat, s: 2 },
        { v: r.fatBirimFiyat, s: 2 },
        { v: r.irsTutar, s: 2 },
        { v: r.fatTutar, s: 2 },
        { v: r.farkTutar, s: rowStyleFor(r.severity, true) },
        { v: r.aciklama, s: 0 },
      ]),
    ];

    const siteRows = [
      ['Şantiye Adı', 'Kalem Sayısı', 'Hatalı Satır', 'Gerçekleşen Maliyet (TL)', 'Fatura Toplamı (TL)', 'Yakalanan Risk (TL)'].map((h) => ({ v: h, s: headerStyle })),
      ...analytics.santiyeler.map((s) => [
        s.name,
        s.kalem,
        s.hata,
        { v: s.tutar, s: 2 },
        { v: s.faturaTutar, s: 2 },
        { v: s.risk, s: s.risk > 0 ? 4 : 2 },
      ]),
    ];

    const blob = await createXlsxBlob([
      { name: 'Özet ve Hatalar', rows: summaryRows },
      { name: 'Tam Karşılaştırma', rows: allRows },
      { name: 'Şantiye Özeti', rows: siteRows },
    ]);

    triggerDownload(blob, 'santiye-irsaliye-fatura-kontrol-raporu.xlsx');
    showToast('Renkli ve özetli Excel raporu indirildi.');
  }

  async function exportHakedisExcel() {
    const res = getHakedisResult();
    const rows = [
      [{ v: 'TAŞERON HAKEDİŞ İCMAL RAPORU', s: 1 }, { v: '', s: 1 }, { v: '', s: 1 }, { v: '', s: 1 }, { v: '', s: 1 }],
      ['Taşeron Firma', $('hkFirma').value],
      ['Şantiye / Proje', $('hkSantiye').value],
      ['Hakediş No / Dönem', $('hkDonem').value],
      ['Tarih', new Date().toLocaleDateString('tr-TR')],
      [],
      ['İş Kalemi', 'Miktar', 'Birim', 'Birim Fiyat (TL)', 'Tutar (TL)'].map((h) => ({ v: h, s: 1 })),
      ...state.hakedisItems.map((it) => [
        it.ad,
        { v: it.miktar, s: 2 },
        it.birim,
        { v: it.fiyat, s: 2 },
        { v: it.miktar * it.fiyat, s: 2 },
      ]),
      [],
      [{ v: 'HAKEDİŞ İCMAL ÖZETİ', s: 9 }, '', '', '', { v: 'TUTAR (TL)', s: 9 }],
      ['Yapılan İş Tutarı (KDV Hariç)', '', '', '', { v: res.brutTutar, s: 10 }],
      [`Hesaplanan KDV (%${res.kdvOrani})`, '', '', '', { v: res.toplamKdv, s: 2 }],
      ['Tevkif Edilen KDV (KDV2)', '', '', '', { v: res.tevkifEdilenKdv, s: 2 }],
      ['Taşerona Kalan KDV', '', '', '', { v: res.taseronaKalanKdv, s: 2 }],
      [{ v: 'Fatura Ödenecek Toplam', s: 9 }, '', '', '', { v: res.faturaOdenecekToplam, s: 10 }],
      [`Teminat Kesintisi (%${res.teminatOrani})`, '', '', '', { v: -res.teminatKesintisi, s: 4 }],
      [`Stopaj Kesintisi (%${res.stopajOrani})`, '', '', '', { v: -res.stopajKesintisi, s: 4 }],
      ['Avans Mahsubu', '', '', '', { v: -res.avansMahsubu, s: 4 }],
      ['Diğer Kesinti / Ceza', '', '', '', { v: -res.digerKesinti, s: 4 }],
      [{ v: 'TAŞERONA ÖDENECEK NET TUTAR', s: 7 }, '', '', '', { v: res.netOdenecek, s: 8 }],
    ];

    const blob = await createXlsxBlob([{ name: 'Hakediş İcmali', rows }]);
    triggerDownload(blob, 'taseron-hakedis-icmal.xlsx');
    showToast('Hakediş Excel dosyası indirildi.');
  }

  async function downloadSampleExcelFiles() {
    const { irsaliyeler, faturalar } = getSampleDatasets();
    const toRows = (arr) => {
      const headers = Object.keys(arr[0]);
      return [
        headers.map((h) => ({ v: h, s: 1 })),
        ...arr.map((obj) =>
          headers.map((h) => {
            const val = obj[h];
            return typeof val === 'number' ? { v: val, s: 2 } : val;
          })
        ),
      ];
    };

    const irsBlob = await createXlsxBlob([{ name: 'İrsaliyeler', rows: toRows(irsaliyeler) }]);
    triggerDownload(irsBlob, 'ornek_santiye_irsaliyeleri.xlsx');

    setTimeout(async () => {
      const fatBlob = await createXlsxBlob([{ name: 'Faturalar', rows: toRows(faturalar) }]);
      triggerDownload(fatBlob, 'ornek_tedarikci_faturalari.xlsx');
      showToast('2 adet örnek Excel dosyası indirildi. Sürükleyip bırakarak test edebilirsiniz!');
    }, 350);
  }

  function updatePrintMeta() {
    const now = new Date().toLocaleString('tr-TR');
    const irsName = state.irsaliyeFile ? state.irsaliyeFile.fileName : '—';
    const fatName = state.faturaFile ? state.faturaFile.fileName : '—';
    $('printMeta').textContent = `Tarih: ${now} • İrsaliye: ${irsName} • Fatura: ${fatName}`;
  }

  // --- Başlatma ---
  function init() {
    initTabs();
    initHakedis();

    setupDropzone('dropIrsaliye', 'fileIrsaliye', 'btnPickIrsaliye', 'irsaliye');
    setupDropzone('dropFatura', 'fileFatura', 'btnPickFatura', 'fatura');

    $('btnLoadDemo').addEventListener('click', loadDemoData);
    $('btnLoadDemoFromAnalytics').addEventListener('click', () => {
      loadDemoData();
      showToast('Örnek veriler analiz sekmesine yüklendi.');
    });
    $('btnRunReconcile').addEventListener('click', () => {
      runReconciliation();
      showToast('Eşleştirme yeni sütun ayarlarıyla güncellendi.');
    });

    $('btnDownloadSampleFiles').addEventListener('click', downloadSampleExcelFiles);
    $('btnExportReconcileXlsx').addEventListener('click', exportReconcileExcel);
    $('btnExportHakedisXlsx').addEventListener('click', exportHakedisExcel);

    $('btnPrintReconcile').addEventListener('click', () => {
      updatePrintMeta();
      window.print();
    });
    $('btnPrintAnalytics').addEventListener('click', () => {
      updatePrintMeta();
      window.print();
    });
    $('btnPrintHakedis').addEventListener('click', () => {
      updatePrintMeta();
      window.print();
    });

    document.querySelectorAll('#filterPills .pill').forEach((pill) => {
      pill.addEventListener('click', () => {
        state.activeFilter = pill.dataset.filter;
        updateFilterPills();
        renderReconcileTable();
      });
    });

    $('siteFilter').addEventListener('change', (e) => {
      state.activeSite = e.target.value;
      renderReconcileTable();
    });

    $('searchInput').addEventListener('input', (e) => {
      state.searchQuery = e.target.value;
      renderReconcileTable();
    });
  }

  document.addEventListener('DOMContentLoaded', init);
})();
