/**
 * reconcile.js
 * İnşaat muhasebesi için İrsaliye-Fatura çapraz kontrol motoru,
 * şantiye/tedarikçi analizleri ve taşeron hakediş hesaplama fonksiyonları.
 */
(function (global) {
  'use strict';

  function trLower(str) {
    return String(str || '').trim().toLocaleLowerCase('tr-TR');
  }

  function normalizeTextKey(str) {
    return trLower(str)
      .replace(/ı/g, 'i')
      .replace(/ğ/g, 'g')
      .replace(/ü/g, 'u')
      .replace(/ş/g, 's')
      .replace(/ö/g, 'o')
      .replace(/ç/g, 'c')
      .replace(/[^a-z0-9]/g, '');
  }

  function normalizeDocNo(str) {
    return String(str || '')
      .trim()
      .toUpperCase()
      .replace(/\s+/g, '');
  }

  /**
   * Türkçe veya İngilizce formatlı sayı metinlerini ("1.250,50", "₺ 24.200,00", "120.5") sayıya çevirir.
   */
  function parseTurkishNumber(val) {
    if (typeof val === 'number') return Number.isFinite(val) ? val : 0;
    if (val === null || val === undefined) return 0;

    let s = String(val).trim();
    if (!s) return 0;

    const match = /-?[\d.,]+/.exec(s);
    if (!match) return 0;
    s = match[0];

    const hasComma = s.includes(',');
    const hasDot = s.includes('.');

    if (hasComma && hasDot) {
      if (s.lastIndexOf(',') > s.lastIndexOf('.')) {
        // 1.250,50 -> 1250.50
        s = s.replace(/\./g, '').replace(',', '.');
      } else {
        // 1,250.50 -> 1250.50
        s = s.replace(/,/g, '');
      }
    } else if (hasComma) {
      const parts = s.split(',');
      if (parts.length > 2) {
        s = s.replace(/,/g, '');
      } else {
        s = s.replace(',', '.');
      }
    } else if (hasDot) {
      const parts = s.split('.');
      if (parts.length > 2) {
        s = s.replace(/\./g, '');
      } else if (parts.length === 2 && parts[1] === '000' && parts[0] !== '0' && parts[0] !== '-0') {
        s = s.replace('.', '');
      }
    }

    const num = Number(s);
    return Number.isFinite(num) ? num : 0;
  }

  /**
   * Excel başlıklarından alan adını akıllı tahmin eder.
   */
  function guessColumnMapping(headers, mode) {
    const aliases = {
      irsaliyeNo: ['irsaliyeno', 'irsno', 'irsaliye', 'belgeno', 'fisno', 'sevkno', 'evrakno'],
      faturaNo: ['faturano', 'fatno', 'faturanumarasi'],
      tarih: ['tarih', 'irstarih', 'irsaliyetarihi', 'faturatarihi', 'belgetarihi', 'date'],
      santiye: ['santiye', 'santiyeadi', 'proje', 'projeadi', 'bolge', 'lokasyon'],
      tedarikci: ['tedarikci', 'firma', 'cari', 'cariadi', 'unvan', 'satici', 'tedarikcifirma'],
      malzeme: ['malzeme', 'urun', 'kalem', 'cins', 'aciklama', 'hizmet', 'malzemecinsi', 'stokadi'],
      birim: ['birim', 'olcubirimi', 'brm', 'unit'],
      miktar: ['miktar', 'adet', 'tonaj', 'metraj', 'qty', 'miktarı'],
      birimFiyat: ['birimfiyat', 'bfiyat', 'fiyat', 'birimfiyati', 'unitprice'],
      tutar: ['tutar', 'toplam', 'toplamtutar', 'nettutar', 'kdvhariçtutar', 'kdvhariç', 'matrah', 'bedel'],
    };

    const normHeaders = headers.map((h) => ({ raw: h, norm: normalizeTextKey(h) }));
    const result = {};

    const fields =
      mode === 'irsaliye'
        ? ['irsaliyeNo', 'tarih', 'santiye', 'tedarikci', 'malzeme', 'birim', 'miktar', 'birimFiyat', 'tutar']
        : ['faturaNo', 'irsaliyeNo', 'tarih', 'santiye', 'tedarikci', 'malzeme', 'birim', 'miktar', 'birimFiyat', 'tutar'];

    const usedHeaders = new Set();

    // 1. Tam eşleşme önceliği
    for (const field of fields) {
      const candidates = aliases[field] || [];
      const found = normHeaders.find((h) => !usedHeaders.has(h.raw) && candidates.includes(h.norm));
      if (found) {
        result[field] = found.raw;
        usedHeaders.add(found.raw);
      }
    }

    // 2. Kısmi (içerir) eşleşme
    for (const field of fields) {
      if (result[field]) continue;
      const candidates = aliases[field] || [];
      const found = normHeaders.find(
        (h) => !usedHeaders.has(h.raw) && candidates.some((c) => h.norm.includes(c) || c.includes(h.norm))
      );
      if (found) {
        result[field] = found.raw;
        usedHeaders.add(found.raw);
      } else {
        result[field] = '';
      }
    }

    return result;
  }

  function normalizeRecord(raw, map, idx) {
    const get = (field) => (map[field] ? raw[map[field]] : '');
    const miktar = parseTurkishNumber(get('miktar'));
    let birimFiyat = parseTurkishNumber(get('birimFiyat'));
    let tutar = parseTurkishNumber(get('tutar'));

    if (tutar === 0 && miktar > 0 && birimFiyat > 0) {
      tutar = Math.round(miktar * birimFiyat * 100) / 100;
    } else if (birimFiyat === 0 && miktar > 0 && tutar > 0) {
      birimFiyat = Math.round((tutar / miktar) * 100) / 100;
    }

    return {
      rowIndex: idx + 1,
      irsaliyeNo: String(get('irsaliyeNo') || '').trim(),
      normIrsaliye: normalizeDocNo(get('irsaliyeNo')),
      faturaNo: String(get('faturaNo') || '').trim(),
      tarih: String(get('tarih') || '').trim(),
      santiye: String(get('santiye') || 'Belirtilmemiş').trim() || 'Belirtilmemiş',
      tedarikci: String(get('tedarikci') || 'Bilinmeyen Tedarikçi').trim() || 'Bilinmeyen Tedarikçi',
      malzeme: String(get('malzeme') || 'Genel Malzeme').trim() || 'Genel Malzeme',
      normMalzeme: normalizeTextKey(get('malzeme')),
      birim: String(get('birim') || '').trim(),
      miktar,
      birimFiyat,
      tutar,
    };
  }

  function categorizeMaterial(malzeme) {
    const m = trLower(malzeme);
    if (/beton|c25|c30|c35|c40|pompa|şap|sap/.test(m)) return 'Hazır Beton & Şap';
    if (/demir|çelik|celik|nervür|hasır|hasir|profil/.test(m)) return 'Demir & Çelik';
    if (/nakliye|hafriyat|kamyon|taşıma|tasima|vinç|vinc/.test(m)) return 'Nakliye & İş Makinesi';
    if (/kum|çakıl|cakil|mıcır|micir|agrega|tuğla|tugla|bims|gazbeton|çimento|cimento/.test(m))
      return 'Agrega & Kaba Malzeme';
    if (/kablo|boru|tesisat|elektrik|armatür|seramik|boya|yalıtım|yalitim|membran/.test(m))
      return 'Mekanik, Elektrik & İnce İş';
    return 'Diğer Malzeme / Hizmet';
  }

  /**
   * İki normalize edilmiş kayıt listesini çapraz kontrol eder.
   */
  function reconcileDocuments(rawIrsaliyeler, rawFaturalar, irsMap, fatMap) {
    const irsList = rawIrsaliyeler
      .map((r, i) => normalizeRecord(r, irsMap, i))
      .filter((r) => r.normIrsaliye !== '' || r.tutar > 0);
    const fatList = rawFaturalar
      .map((r, i) => normalizeRecord(r, fatMap, i))
      .filter((r) => r.normIrsaliye !== '' || r.tutar > 0 || r.faturaNo !== '');

    // İrsaliyeleri normIrsaliye'ye göre grupla
    const irsGroups = new Map();
    for (const irs of irsList) {
      if (!irs.normIrsaliye) continue;
      if (!irsGroups.has(irs.normIrsaliye)) irsGroups.set(irs.normIrsaliye, []);
      irsGroups.get(irs.normIrsaliye).push({ item: irs, matched: false });
    }

    const results = [];
    const seenInvoiceKeys = new Set();

    for (const fat of fatList) {
      if (!fat.normIrsaliye) {
        results.push({
          durum: 'İRSALİYE YOK',
          severity: 'danger',
          irsaliyeNo: '—',
          faturaNo: fat.faturaNo || `Satır #${fat.rowIndex}`,
          tarih: fat.tarih,
          santiye: fat.santiye,
          tedarikci: fat.tedarikci,
          malzeme: fat.malzeme,
          birim: fat.birim,
          irsMiktar: 0,
          fatMiktar: fat.miktar,
          irsBirimFiyat: 0,
          fatBirimFiyat: fat.birimFiyat,
          irsTutar: 0,
          fatTutar: fat.tutar,
          farkTutar: fat.tutar,
          riskTutari: fat.tutar,
          aciklama: 'Faturada irsaliye numarası belirtilmemiş; şantiye irsaliyesi ile doğrulanamadı.',
        });
        continue;
      }

      const candidates = irsGroups.get(fat.normIrsaliye);
      if (!candidates || candidates.length === 0) {
        results.push({
          durum: 'İRSALİYE YOK',
          severity: 'danger',
          irsaliyeNo: fat.irsaliyeNo,
          faturaNo: fat.faturaNo || '—',
          tarih: fat.tarih,
          santiye: fat.santiye,
          tedarikci: fat.tedarikci,
          malzeme: fat.malzeme,
          birim: fat.birim,
          irsMiktar: 0,
          fatMiktar: fat.miktar,
          irsBirimFiyat: 0,
          fatBirimFiyat: fat.birimFiyat,
          irsTutar: 0,
          fatTutar: fat.tutar,
          farkTutar: fat.tutar,
          riskTutari: fat.tutar,
          aciklama: `Şantiye irsaliye listesinde "${fat.irsaliyeNo}" numaralı teslimat bulunamadı.`,
        });
        continue;
      }

      // Aynı (irsaliye + malzeme) faturada birden fazla kez geçiyorsa mükerrer yakala
      const dupKey = `${fat.normIrsaliye}::${fat.normMalzeme || 'genel'}`;
      let candidate =
        candidates.find((c) => !c.matched && c.item.normMalzeme === fat.normMalzeme) ||
        candidates.find((c) => !c.matched && (c.item.normMalzeme.includes(fat.normMalzeme) || fat.normMalzeme.includes(c.item.normMalzeme))) ||
        candidates.find((c) => !c.matched);

      if (!candidate) {
        const isDuplicate = seenInvoiceKeys.has(dupKey);
        results.push({
          durum: isDuplicate ? 'MÜKERRER FATURA' : 'İRSALİYE YOK',
          severity: 'danger',
          irsaliyeNo: fat.irsaliyeNo,
          faturaNo: fat.faturaNo || '—',
          tarih: fat.tarih,
          santiye: fat.santiye,
          tedarikci: fat.tedarikci,
          malzeme: fat.malzeme,
          birim: fat.birim,
          irsMiktar: 0,
          fatMiktar: fat.miktar,
          irsBirimFiyat: 0,
          fatBirimFiyat: fat.birimFiyat,
          irsTutar: 0,
          fatTutar: fat.tutar,
          farkTutar: fat.tutar,
          riskTutari: fat.tutar,
          aciklama: isDuplicate
            ? `Aynı irsaliye (${fat.irsaliyeNo}) ve kalem bu fatura listesinde mükerrer (çift) faturalandırılmış.`
            : `"${fat.irsaliyeNo}" irsaliyesindeki kalemler eşleşti ancak faturada irsaliyede olmayan ekstra kalem ("${fat.malzeme}") bulunuyor.`,
        });
        continue;
      }

      candidate.matched = true;
      seenInvoiceKeys.add(dupKey);

      const irs = candidate.item;
      const qtyDiff = Math.round((fat.miktar - irs.miktar) * 1000) / 1000;
      const priceDiff = Math.round((fat.birimFiyat - irs.birimFiyat) * 100) / 100;
      const amtDiff = Math.round((fat.tutar - irs.tutar) * 100) / 100;

      const hasQtyIssue = (irs.miktar > 0 || fat.miktar > 0) && Math.abs(qtyDiff) > 0.01;
      const hasPriceIssue = (irs.birimFiyat > 0 || fat.birimFiyat > 0) && Math.abs(priceDiff) > 0.05;
      const hasAmtIssue = Math.abs(amtDiff) > 1.0;

      let durum = 'EŞLEŞTİ';
      let severity = 'ok';
      let aciklama = 'Miktar ve birim fiyat şantiye irsaliyesi ile birebir uyumlu.';

      if (hasQtyIssue && hasPriceIssue) {
        durum = 'MİKTAR + FİYAT FARKI';
        severity = 'danger';
        aciklama = `Miktar farkı (${qtyDiff > 0 ? '+' : ''}${qtyDiff} ${irs.birim || fat.birim}) ve birim fiyat farkı (${priceDiff > 0 ? '+' : ''}${priceDiff.toLocaleString('tr-TR')} TL) tespit edildi.`;
      } else if (hasQtyIssue) {
        durum = 'MİKTAR FARKI';
        severity = qtyDiff > 0 ? 'danger' : 'warn';
        aciklama = `Şantiyeye ${irs.miktar} ${irs.birim || fat.birim} girmiş, faturada ${fat.miktar} ${fat.birim || irs.birim} kesilmiş (${qtyDiff > 0 ? '+' : ''}${qtyDiff} fark).`;
      } else if (hasPriceIssue) {
        durum = 'FİYAT FARKI';
        severity = priceDiff > 0 ? 'danger' : 'warn';
        aciklama = `Sözleşme/irsaliye birim fiyatı ${irs.birimFiyat.toLocaleString('tr-TR')} TL iken fatura ${fat.birimFiyat.toLocaleString('tr-TR')} TL üzerinden kesilmiş.`;
      } else if (hasAmtIssue) {
        durum = 'TUTAR FARKI';
        severity = amtDiff > 0 ? 'danger' : 'warn';
        aciklama = `Satır toplam tutarı irsaliye hesabından ${amtDiff > 0 ? '+' : ''}${amtDiff.toLocaleString('tr-TR')} TL sapma gösteriyor.`;
      }

      results.push({
        durum,
        severity,
        irsaliyeNo: irs.irsaliyeNo || fat.irsaliyeNo,
        faturaNo: fat.faturaNo || '—',
        tarih: fat.tarih || irs.tarih,
        santiye: irs.santiye !== 'Belirtilmemiş' ? irs.santiye : fat.santiye,
        tedarikci: fat.tedarikci !== 'Bilinmeyen Tedarikçi' ? fat.tedarikci : irs.tedarikci,
        malzeme: irs.malzeme || fat.malzeme,
        birim: irs.birim || fat.birim,
        irsMiktar: irs.miktar,
        fatMiktar: fat.miktar,
        irsBirimFiyat: irs.birimFiyat,
        fatBirimFiyat: fat.birimFiyat,
        irsTutar: irs.tutar,
        fatTutar: fat.tutar,
        farkTutar: amtDiff,
        riskTutari: amtDiff > 0 ? amtDiff : 0,
        aciklama,
      });
    }

    // Faturası henüz gelmemiş şantiye irsaliyeleri
    for (const group of irsGroups.values()) {
      for (const entry of group) {
        if (!entry.matched) {
          const irs = entry.item;
          results.push({
            durum: 'FATURA BEKLİYOR',
            severity: 'info',
            irsaliyeNo: irs.irsaliyeNo,
            faturaNo: 'Henüz Gelmedi',
            tarih: irs.tarih,
            santiye: irs.santiye,
            tedarikci: irs.tedarikci,
            malzeme: irs.malzeme,
            birim: irs.birim,
            irsMiktar: irs.miktar,
            fatMiktar: 0,
            irsBirimFiyat: irs.birimFiyat,
            fatBirimFiyat: 0,
            irsTutar: irs.tutar,
            fatTutar: 0,
            farkTutar: -irs.tutar,
            riskTutari: 0,
            aciklama: 'Malzeme şantiyeye teslim alınmış ancak tedarikçi faturası henüz listede yok.',
          });
        }
      }
    }

    // Özet KPI ve analitik verileri hesapla
    const summary = {
      toplamKalem: results.length,
      eslesenSayisi: 0,
      hataliSayisi: 0,
      bekleyenSayisi: 0,
      toplamIrsaliyeTutar: 0,
      toplamFaturaTutar: 0,
      fazlaFaturalananRisk: 0,
      bekleyenIrsaliyeTutar: 0,
    };

    const santiyeMap = new Map();
    const kategoriMap = new Map();
    const tedarikciMap = new Map();

    for (const row of results) {
      summary.toplamIrsaliyeTutar += row.irsTutar;
      summary.toplamFaturaTutar += row.fatTutar;
      summary.fazlaFaturalananRisk += row.riskTutari;

      if (row.durum === 'EŞLEŞTİ') {
        summary.eslesenSayisi++;
      } else if (row.durum === 'FATURA BEKLİYOR') {
        summary.bekleyenSayisi++;
        summary.bekleyenIrsaliyeTutar += row.irsTutar;
      } else {
        summary.hataliSayisi++;
      }

      // Maliyet dağılımında gerçekleşmiş tutar (irsaliye varsa baz tutar, yoksa fatura tutarı)
      const costBasis = row.irsTutar > 0 ? row.irsTutar : row.fatTutar;

      const sRec = santiyeMap.get(row.santiye) || { name: row.santiye, tutar: 0, faturaTutar: 0, risk: 0, kalem: 0, hata: 0 };
      sRec.tutar += costBasis;
      sRec.faturaTutar += row.fatTutar;
      sRec.risk += row.riskTutari;
      sRec.kalem++;
      if (row.durum !== 'EŞLEŞTİ' && row.durum !== 'FATURA BEKLİYOR') sRec.hata++;
      santiyeMap.set(row.santiye, sRec);

      const cat = categorizeMaterial(row.malzeme);
      const kRec = kategoriMap.get(cat) || { name: cat, tutar: 0, kalem: 0 };
      kRec.tutar += costBasis;
      kRec.kalem++;
      kategoriMap.set(cat, kRec);

      const tRec = tedarikciMap.get(row.tedarikci) || { name: row.tedarikci, faturaTutar: 0, irsaliyeTutar: 0, risk: 0, kalem: 0, hata: 0 };
      tRec.faturaTutar += row.fatTutar;
      tRec.irsaliyeTutar += row.irsTutar;
      tRec.risk += row.riskTutari;
      tRec.kalem++;
      if (row.durum !== 'EŞLEŞTİ' && row.durum !== 'FATURA BEKLİYOR') tRec.hata++;
      tedarikciMap.set(row.tedarikci, tRec);
    }

    const byAmountDesc = (a, b) => (b.tutar || b.faturaTutar) - (a.tutar || a.faturaTutar);

    return {
      rows: results,
      summary,
      analytics: {
        santiyeler: Array.from(santiyeMap.values()).sort(byAmountDesc),
        kategoriler: Array.from(kategoriMap.values()).sort(byAmountDesc),
        tedarikciler: Array.from(tedarikciMap.values()).sort((a, b) => b.faturaTutar - a.faturaTutar),
      },
    };
  }

  /**
   * Taşeron Hakediş & Kesinti Hesaplayıcı
   */
  function calculateHakedis(input) {
    const brutTutar = Math.max(0, Number(input.brutTutar) || 0);
    const kdvOrani = Math.max(0, Number(input.kdvOrani) || 0);
    const tevkifatPay = Math.max(0, Math.min(1, Number(input.tevkifatPay) || 0));
    const teminatOrani = Math.max(0, Number(input.teminatOrani) || 0);
    const stopajOrani = Math.max(0, Number(input.stopajOrani) || 0);
    const avansMahsubu = Math.max(0, Number(input.avansMahsubu) || 0);
    const digerKesinti = Math.max(0, Number(input.digerKesinti) || 0);

    const toplamKdv = brutTutar * (kdvOrani / 100);
    const tevkifEdilenKdv = toplamKdv * tevkifatPay;
    const taseronaKalanKdv = toplamKdv - tevkifEdilenKdv;

    const kdvDahilGenelToplam = brutTutar + toplamKdv;
    const faturaOdenecekToplam = brutTutar + taseronaKalanKdv;

    const teminatKesintisi = brutTutar * (teminatOrani / 100);
    const stopajKesintisi = brutTutar * (stopajOrani / 100);
    const toplamKesinti = teminatKesintisi + stopajKesintisi + avansMahsubu + digerKesinti;
    const netOdenecek = faturaOdenecekToplam - toplamKesinti;

    return {
      brutTutar,
      kdvOrani,
      toplamKdv,
      tevkifatPay,
      tevkifEdilenKdv,
      taseronaKalanKdv,
      kdvDahilGenelToplam,
      faturaOdenecekToplam,
      teminatOrani,
      teminatKesintisi,
      stopajOrani,
      stopajKesintisi,
      avansMahsubu,
      digerKesinti,
      toplamKesinti,
      netOdenecek,
    };
  }

  const api = {
    parseTurkishNumber,
    guessColumnMapping,
    reconcileDocuments,
    calculateHakedis,
    categorizeMaterial,
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  } else {
    global.ReconcileEngine = api;
  }
})(typeof window !== 'undefined' ? window : globalThis);
