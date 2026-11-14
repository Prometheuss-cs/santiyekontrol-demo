/**
 * xlsx-lite.js
 * Sıfır bağımlılıklı, tarayıcıda %100 çevrimdışı çalışan Excel (.xlsx) ve CSV okuyucu / yazıcı.
 * Harici kütüphane veya sunucu gerektirmez; veriler cihaz dışına çıkmaz.
 */
(function (global) {
  'use strict';

  const textEncoder = new TextEncoder();
  const textDecoderUtf8 = new TextDecoder('utf-8');

  // --- CRC-32 Tablosu (ZIP arşivi oluşturmak için) ---
  const CRC_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let k = 0; k < 8; k++) {
        c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
      }
      table[i] = c >>> 0;
    }
    return table;
  })();

  function crc32(bytes) {
    let crc = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) {
      crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  // --- Deflate / Inflate yardımcıları (Tarayıcı yerel API'si) ---
  async function inflateRaw(compressedBytes) {
    const stream = new Blob([compressedBytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    const buf = await new Response(stream).arrayBuffer();
    return new Uint8Array(buf);
  }

  async function deflateRaw(rawBytes) {
    if (typeof CompressionStream === 'undefined') {
      return { method: 0, data: rawBytes };
    }
    try {
      const stream = new Blob([rawBytes]).stream().pipeThrough(new CompressionStream('deflate-raw'));
      const buf = await new Response(stream).arrayBuffer();
      return { method: 8, data: new Uint8Array(buf) };
    } catch (_) {
      return { method: 0, data: rawBytes };
    }
  }

  // --- ZIP Okuyucu (Central Directory tabanlı) ---
  async function unzipEntries(arrayBuffer) {
    const bytes = new Uint8Array(arrayBuffer);
    const view = new DataView(arrayBuffer);

    // End of Central Directory (EOCD) imzasını sondan ara: 0x06054b50
    const minOffset = Math.max(0, bytes.length - 65557);
    let eocdOffset = -1;
    for (let i = bytes.length - 22; i >= minOffset; i--) {
      if (bytes[i] === 0x50 && bytes[i + 1] === 0x4b && bytes[i + 2] === 0x05 && bytes[i + 3] === 0x06) {
        eocdOffset = i;
        break;
      }
    }
    if (eocdOffset === -1) {
      throw new Error('Geçerli bir Excel (.xlsx / ZIP) dosyası değil.');
    }

    const totalEntries = view.getUint16(eocdOffset + 10, true);
    const cdOffset = view.getUint32(eocdOffset + 16, true);

    const files = new Map();
    let ptr = cdOffset;

    for (let idx = 0; idx < totalEntries; idx++) {
      if (view.getUint32(ptr, true) !== 0x02014b50) {
        break;
      }
      const method = view.getUint16(ptr + 10, true);
      const compSize = view.getUint32(ptr + 20, true);
      const nameLen = view.getUint16(ptr + 28, true);
      const extraLen = view.getUint16(ptr + 30, true);
      const commentLen = view.getUint16(ptr + 32, true);
      const localHeaderOffset = view.getUint32(ptr + 42, true);

      const name = textDecoderUtf8.decode(bytes.subarray(ptr + 46, ptr + 46 + nameLen));
      ptr += 46 + nameLen + extraLen + commentLen;

      if (name.endsWith('/')) continue;

      if (view.getUint32(localHeaderOffset, true) !== 0x04034b50) {
        continue;
      }
      const lhNameLen = view.getUint16(localHeaderOffset + 26, true);
      const lhExtraLen = view.getUint16(localHeaderOffset + 28, true);
      const dataStart = localHeaderOffset + 30 + lhNameLen + lhExtraLen;
      const compressed = bytes.subarray(dataStart, dataStart + compSize);

      let uncompressed;
      if (method === 0) {
        uncompressed = compressed;
      } else if (method === 8) {
        uncompressed = await inflateRaw(compressed);
      } else {
        throw new Error(`Desteklenmeyen ZIP sıkıştırma yöntemi (${method}): ${name}`);
      }
      files.set(name, uncompressed);
    }

    return files;
  }

  // --- ZIP Yazıcı ---
  async function createZipBlob(entries) {
    const localParts = [];
    const centralParts = [];
    let offset = 0;

    for (const entry of entries) {
      const nameBytes = textEncoder.encode(entry.name);
      const rawBytes = typeof entry.content === 'string' ? textEncoder.encode(entry.content) : entry.content;
      const crc = crc32(rawBytes);
      const { method, data: compBytes } = await deflateRaw(rawBytes);

      const lh = new Uint8Array(30 + nameBytes.length);
      const lhv = new DataView(lh.buffer);
      lhv.setUint32(0, 0x04034b50, true);
      lhv.setUint16(4, 20, true);
      lhv.setUint16(6, 0x0800, true); // UTF-8 flag
      lhv.setUint16(8, method, true);
      lhv.setUint16(10, 0, true);
      lhv.setUint16(12, 0x5421, true);
      lhv.setUint32(14, crc, true);
      lhv.setUint32(18, compBytes.length, true);
      lhv.setUint32(22, rawBytes.length, true);
      lhv.setUint16(26, nameBytes.length, true);
      lhv.setUint16(28, 0, true);
      lh.set(nameBytes, 30);

      localParts.push(lh, compBytes);

      const cd = new Uint8Array(46 + nameBytes.length);
      const cdv = new DataView(cd.buffer);
      cdv.setUint32(0, 0x02014b50, true);
      cdv.setUint16(4, 20, true);
      cdv.setUint16(6, 20, true);
      cdv.setUint16(8, 0x0800, true);
      cdv.setUint16(10, method, true);
      cdv.setUint16(12, 0, true);
      cdv.setUint16(14, 0x5421, true);
      cdv.setUint32(16, crc, true);
      cdv.setUint32(20, compBytes.length, true);
      cdv.setUint32(24, rawBytes.length, true);
      cdv.setUint16(28, nameBytes.length, true);
      cdv.setUint16(30, 0, true);
      cdv.setUint16(32, 0, true);
      cdv.setUint16(34, 0, true);
      cdv.setUint16(36, 0, true);
      cdv.setUint32(38, 0, true);
      cdv.setUint32(42, offset, true);
      cd.set(nameBytes, 46);

      centralParts.push(cd);
      offset += lh.length + compBytes.length;
    }

    let cdSize = 0;
    for (const part of centralParts) cdSize += part.length;

    const eocd = new Uint8Array(22);
    const ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(4, 0, true);
    ev.setUint16(6, 0, true);
    ev.setUint16(8, entries.length, true);
    ev.setUint16(10, entries.length, true);
    ev.setUint32(12, cdSize, true);
    ev.setUint32(16, offset, true);
    ev.setUint16(20, 0, true);

    return new Blob([...localParts, ...centralParts, eocd], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  }

  // --- XML Yardımcıları ---
  function parseXml(bytes) {
    if (!bytes) return null;
    const text = textDecoderUtf8.decode(bytes);
    return new DOMParser().parseFromString(text, 'application/xml');
  }

  function colLettersToIndex(letters) {
    let idx = 0;
    for (let i = 0; i < letters.length; i++) {
      idx = idx * 26 + (letters.charCodeAt(i) - 64);
    }
    return idx - 1;
  }

  function indexToColLetters(idx) {
    let n = idx + 1;
    let out = '';
    while (n > 0) {
      const rem = (n - 1) % 26;
      out = String.fromCharCode(65 + rem) + out;
      n = Math.floor((n - 1) / 26);
    }
    return out;
  }

  function normalizeZipPath(baseDir, target) {
    if (!target) return '';
    if (target.startsWith('/')) return target.slice(1);
    const parts = (baseDir ? baseDir.split('/') : []).concat(target.split('/'));
    const stack = [];
    for (const p of parts) {
      if (!p || p === '.') continue;
      if (p === '..') stack.pop();
      else stack.push(p);
    }
    return stack.join('/');
  }

  function isDateFormatCode(code) {
    if (!code) return false;
    const cleaned = code
      .replace(/"[^"]*"/g, '')
      .replace(/\[[^\]]*\]/g, '')
      .toLowerCase();
    return /[yd]/.test(cleaned) || (/m/.test(cleaned) && /[\/\-.]/.test(cleaned));
  }

  function excelSerialToTurkishDate(serial) {
    if (typeof serial !== 'number' || !Number.isFinite(serial) || serial < 1) {
      return String(serial);
    }
    // Excel 1900 artık yıl hatası düzeltmesi (serial 60 = var olmayan 29.02.1900)
    const adjusted = serial > 59 ? serial - 1 : serial;
    const utcMs = Date.UTC(1899, 11, 31) + Math.round(adjusted) * 86400000;
    const d = new Date(utcMs);
    const dd = String(d.getUTCDate()).padStart(2, '0');
    const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
    const yyyy = d.getUTCFullYear();
    return `${dd}.${mm}.${yyyy}`;
  }

  function parseDateStyles(stylesDoc) {
    const dateStyleIndices = new Set();
    if (!stylesDoc) return dateStyleIndices;

    const builtinDateIds = new Set([
      14, 15, 16, 17, 18, 19, 20, 21, 22,
      27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
      45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58,
    ]);

    const customDateIds = new Set();
    const numFmts = stylesDoc.getElementsByTagName('numFmt');
    for (let i = 0; i < numFmts.length; i++) {
      const id = Number(numFmts[i].getAttribute('numFmtId'));
      const code = numFmts[i].getAttribute('formatCode') || '';
      if (isDateFormatCode(code)) {
        customDateIds.add(id);
      }
    }

    const cellXfs = stylesDoc.getElementsByTagName('cellXfs')[0];
    if (cellXfs) {
      const xfs = cellXfs.getElementsByTagName('xf');
      for (let i = 0; i < xfs.length; i++) {
        const fmtId = Number(xfs[i].getAttribute('numFmtId') || 0);
        if (builtinDateIds.has(fmtId) || customDateIds.has(fmtId)) {
          dateStyleIndices.add(i);
        }
      }
    }
    return dateStyleIndices;
  }

  function parseSharedStrings(ssDoc) {
    const list = [];
    if (!ssDoc) return list;
    const siNodes = ssDoc.getElementsByTagName('si');
    for (let i = 0; i < siNodes.length; i++) {
      const si = siNodes[i];
      let str = '';
      // Direkt <t> veya zengin metin <r><t> çocuklarını topla (<rPh> fonetik bloklarını hariç tut)
      for (let j = 0; j < si.childNodes.length; j++) {
        const child = si.childNodes[j];
        if (child.nodeType !== 1) continue;
        const name = child.localName || child.nodeName;
        if (name === 't') {
          str += child.textContent || '';
        } else if (name === 'r') {
          const ts = child.getElementsByTagName('t');
          for (let k = 0; k < ts.length; k++) {
            str += ts[k].textContent || '';
          }
        }
      }
      list.push(str);
    }
    return list;
  }

  function extractSheetRows(sheetDoc, sharedStrings, dateStyleIndices) {
    const rows = [];
    const rowNodes = sheetDoc.getElementsByTagName('row');
    let maxCols = 0;

    for (let i = 0; i < rowNodes.length; i++) {
      const rowNode = rowNodes[i];
      const cells = rowNode.getElementsByTagName('c');
      const rowArr = [];
      let fallbackCol = 0;

      for (let j = 0; j < cells.length; j++) {
        const c = cells[j];
        const ref = c.getAttribute('r');
        let colIdx = fallbackCol;
        if (ref) {
          const match = /^([A-Z]+)/i.exec(ref);
          if (match) {
            colIdx = colLettersToIndex(match[1].toUpperCase());
          }
        }
        fallbackCol = colIdx + 1;

        const type = c.getAttribute('t') || 'n';
        const styleIdx = c.getAttribute('s') !== null ? Number(c.getAttribute('s')) : -1;
        let val = '';

        if (type === 'inlineStr') {
          const isNode = c.getElementsByTagName('is')[0];
          if (isNode) {
            const ts = isNode.getElementsByTagName('t');
            for (let k = 0; k < ts.length; k++) val += ts[k].textContent || '';
          }
        } else {
          const vNode = c.getElementsByTagName('v')[0];
          const raw = vNode ? (vNode.textContent || '') : '';
          if (type === 's') {
            const idx = Number(raw);
            val = sharedStrings[idx] !== undefined ? sharedStrings[idx] : '';
          } else if (type === 'b') {
            val = raw === '1' ? 'DOĞRU' : 'YANLIŞ';
          } else if (type === 'str' || type === 'e') {
            val = raw;
          } else {
            // Sayısal değer veya tarih seri numarası
            if (raw !== '' && !Number.isNaN(Number(raw))) {
              const num = Number(raw);
              val = dateStyleIndices.has(styleIdx) ? excelSerialToTurkishDate(num) : num;
            } else {
              val = raw;
            }
          }
        }

        rowArr[colIdx] = val;
        if (colIdx + 1 > maxCols) maxCols = colIdx + 1;
      }

      // Boş satırları atlama (indeks koru, sonra trim yapacağız)
      const hasContent = rowArr.some((cell) => cell !== undefined && cell !== null && String(cell).trim() !== '');
      if (hasContent) {
        rows.push(rowArr);
      }
    }

    // Dizi boşluklarını '' ile doldur
    return rows.map((r) => {
      const filled = new Array(maxCols);
      for (let c = 0; c < maxCols; c++) {
        filled[c] = r[c] !== undefined && r[c] !== null ? r[c] : '';
      }
      return filled;
    });
  }

  /**
   * Bir .xlsx ArrayBuffer'ını okur ve ilk (veya tüm) sayfaları satır/sütun dizisi olarak döndürür.
   */
  async function readXlsx(arrayBuffer) {
    const files = await unzipEntries(arrayBuffer);
    const workbookDoc = parseXml(files.get('xl/workbook.xml'));
    const relsDoc = parseXml(files.get('xl/_rels/workbook.xml.rels'));
    const ssDoc = parseXml(files.get('xl/sharedStrings.xml'));
    const stylesDoc = parseXml(files.get('xl/styles.xml'));

    const sharedStrings = parseSharedStrings(ssDoc);
    const dateStyleIndices = parseDateStyles(stylesDoc);

    const relMap = new Map();
    if (relsDoc) {
      const relNodes = relsDoc.getElementsByTagName('Relationship');
      for (let i = 0; i < relNodes.length; i++) {
        const r = relNodes[i];
        relMap.set(r.getAttribute('Id'), normalizeZipPath('xl', r.getAttribute('Target')));
      }
    }

    const sheets = [];
    if (workbookDoc) {
      const sheetNodes = workbookDoc.getElementsByTagName('sheet');
      for (let i = 0; i < sheetNodes.length; i++) {
        const s = sheetNodes[i];
        const name = s.getAttribute('name') || `Sayfa${i + 1}`;
        const rId =
          s.getAttribute('r:id') ||
          s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
        const targetPath = relMap.get(rId) || `xl/worksheets/sheet${i + 1}.xml`;
        const sheetBytes = files.get(targetPath);
        if (!sheetBytes) continue;
        const sheetDoc = parseXml(sheetBytes);
        const rows = extractSheetRows(sheetDoc, sharedStrings, dateStyleIndices);
        sheets.push({ name, rows });
      }
    }

    if (sheets.length === 0) {
      // Yedek: doğrudan sheet1.xml ara
      const sheet1 = files.get('xl/worksheets/sheet1.xml');
      if (!sheet1) throw new Error('Excel dosyasında okunabilir çalışma sayfası bulunamadı.');
      sheets.push({
        name: 'Sayfa1',
        rows: extractSheetRows(parseXml(sheet1), sharedStrings, dateStyleIndices),
      });
    }

    return sheets;
  }

  /**
   * CSV / TSV metnini okur (Türkçe Excel'in ';' ayıracını ve Windows-1254 karakter setini destekler).
   */
  function readCsvBuffer(arrayBuffer) {
    let text = textDecoderUtf8.decode(arrayBuffer);
    if (text.includes('\uFFFD')) {
      try {
        text = new TextDecoder('windows-1254').decode(arrayBuffer);
      } catch (_) {
        // UTF-8 ile devam et
      }
    }
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);

    const firstLine = (text.split(/\r?\n/)[0] || '');
    const semiCount = (firstLine.match(/;/g) || []).length;
    const commaCount = (firstLine.match(/,/g) || []).length;
    const tabCount = (firstLine.match(/\t/g) || []).length;
    let delimiter = ',';
    if (tabCount > semiCount && tabCount > commaCount) delimiter = '\t';
    else if (semiCount >= commaCount && semiCount > 0) delimiter = ';';

    const rows = [];
    let currentRow = [];
    let currentVal = '';
    let inQuotes = false;

    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (inQuotes) {
        if (ch === '"') {
          if (text[i + 1] === '"') {
            currentVal += '"';
            i++;
          } else {
            inQuotes = false;
          }
        } else {
          currentVal += ch;
        }
      } else if (ch === '"') {
        inQuotes = true;
      } else if (ch === delimiter) {
        currentRow.push(currentVal.trim());
        currentVal = '';
      } else if (ch === '\r' || ch === '\n') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        currentRow.push(currentVal.trim());
        if (currentRow.some((c) => c !== '')) rows.push(currentRow);
        currentRow = [];
        currentVal = '';
      } else {
        currentVal += ch;
      }
    }
    currentRow.push(currentVal.trim());
    if (currentRow.some((c) => c !== '')) rows.push(currentRow);

    return [{ name: 'CSV', rows }];
  }

  /**
   * Yüklenen File nesnesini (.xlsx / .csv / .tsv) okuyup { headers, records, sheetName } döndürür.
   */
  async function parseSpreadsheetFile(file) {
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    const buf = await file.arrayBuffer();

    let sheets;
    if (ext === 'csv' || ext === 'tsv' || ext === 'txt') {
      sheets = readCsvBuffer(buf);
    } else if (ext === 'xlsx') {
      sheets = await readXlsx(buf);
    } else {
      throw new Error(`Desteklenmeyen dosya türü (.${ext}). Lütfen .xlsx veya .csv yükleyin.`);
    }

    // Satır içeren ilk sayfayı seç
    const targetSheet = sheets.find((s) => s.rows && s.rows.length >= 2) || sheets[0];
    if (!targetSheet || targetSheet.rows.length === 0) {
      throw new Error('Dosya boş görünüyor.');
    }

    // Başlık satırını bul (birden fazla dolu hücre içeren ilk satır)
    let headerIdx = 0;
    for (let i = 0; i < Math.min(5, targetSheet.rows.length); i++) {
      const nonEmpty = targetSheet.rows[i].filter((c) => String(c).trim() !== '').length;
      if (nonEmpty >= 2) {
        headerIdx = i;
        break;
      }
    }

    const rawHeaders = targetSheet.rows[headerIdx].map((h, i) => {
      const clean = String(h || '').trim();
      return clean || `Sütun ${i + 1}`;
    });

    const records = [];
    for (let r = headerIdx + 1; r < targetSheet.rows.length; r++) {
      const row = targetSheet.rows[r];
      const obj = {};
      let hasData = false;
      for (let c = 0; c < rawHeaders.length; c++) {
        const val = row[c] !== undefined ? row[c] : '';
        if (String(val).trim() !== '') hasData = true;
        obj[rawHeaders[c]] = val;
      }
      if (hasData) records.push(obj);
    }

    return {
      fileName: file.name,
      sheetName: targetSheet.name,
      sheetNames: sheets.map((s) => s.name),
      headers: rawHeaders,
      records,
    };
  }

  // --- XLSX Yazıcı (Çok sayfalı, renkli & formatlı gerçek Excel üretimi) ---
  function xmlEscape(val) {
    return String(val === null || val === undefined ? '' : val)
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function buildStylesXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0.00"/>
  </numFmts>
  <fonts count="6">
    <font><sz val="10"/><color rgb="FF0F172A"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FF991B1B"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FF92400E"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FF065F46"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF0F172A"/><name val="Calibri"/></font>
  </fonts>
  <fills count="6">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1E293B"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFEE2E2"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFEF3C7"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD1FAE5"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFE2E8F0"/></left>
      <right style="thin"><color rgb="FFE2E8F0"/></right>
      <top style="thin"><color rgb="FFE2E8F0"/></top>
      <bottom style="thin"><color rgb="FFE2E8F0"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="11">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1"><alignment vertical="center"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="2" fillId="3" borderId="1" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="3" fillId="4" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="3" fillId="4" borderId="1" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="4" fillId="5" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="4" fillId="5" borderId="1" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="5" fillId="0" borderId="1" applyFont="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="5" fillId="0" borderId="1" applyNumberFormat="1" applyFont="1" applyBorder="1"/>
  </cellXfs>
</styleSheet>`;
  }

  /**
   * sheets: Array<{ name: string, columns?: number[], rows: Array<Array<{ v: any, s?: number } | any>> }>
   */
  async function createXlsxBlob(sheets) {
    const zipEntries = [
      {
        name: '_rels/.rels',
        content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
      },
      {
        name: 'xl/styles.xml',
        content: buildStylesXml(),
      },
    ];

    const contentTypeOverrides = [
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ];
    const workbookSheetsXml = [];
    const workbookRelsXml = [];

    sheets.forEach((sheet, idx) => {
      const sheetNum = idx + 1;
      const safeName = xmlEscape((sheet.name || `Sayfa${sheetNum}`).slice(0, 31).replace(/[\\/*?:[\]]/g, ' '));
      contentTypeOverrides.push(
        `<Override PartName="/xl/worksheets/sheet${sheetNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
      );
      workbookSheetsXml.push(`<sheet name="${safeName}" sheetId="${sheetNum}" r:id="rId${sheetNum}"/>`);
      workbookRelsXml.push(
        `<Relationship Id="rId${sheetNum}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${sheetNum}.xml"/>`
      );

      // Sütun genişliklerini otomatik hesapla
      const colWidths = [];
      for (const row of sheet.rows) {
        row.forEach((cell, cIdx) => {
          const val = cell && typeof cell === 'object' && 'v' in cell ? cell.v : cell;
          const len = String(val === null || val === undefined ? '' : val).length;
          colWidths[cIdx] = Math.min(44, Math.max(colWidths[cIdx] || 10, len + 4));
        });
      }
      if (sheet.columns) {
        sheet.columns.forEach((w, cIdx) => {
          if (w) colWidths[cIdx] = w;
        });
      }

      const colsXml = colWidths.length
        ? `<cols>${colWidths.map((w, c) => `<col min="${c + 1}" max="${c + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>`
        : '';

      const rowsXml = sheet.rows
        .map((row, rIdx) => {
          const rNum = rIdx + 1;
          const cellsXml = row
            .map((cell, cIdx) => {
              const isObj = cell !== null && typeof cell === 'object' && 'v' in cell;
              const rawVal = isObj ? cell.v : cell;
              const styleId = isObj && cell.s !== undefined ? cell.s : 0;
              const ref = `${indexToColLetters(cIdx)}${rNum}`;

              if (typeof rawVal === 'number' && Number.isFinite(rawVal)) {
                return `<c r="${ref}" s="${styleId}"><v>${rawVal}</v></c>`;
              }
              return `<c r="${ref}" t="inlineStr" s="${styleId}"><is><t>${xmlEscape(rawVal)}</t></is></c>`;
            })
            .join('');
          return `<row r="${rNum}">${cellsXml}</row>`;
        })
        .join('');

      const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  ${colsXml}
  <sheetData>${rowsXml}</sheetData>
</worksheet>`;

      zipEntries.push({
        name: `xl/worksheets/sheet${sheetNum}.xml`,
        content: sheetXml,
      });
    });

    workbookRelsXml.push(
      `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`
    );

    zipEntries.push({
      name: '[Content_Types].xml',
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-officedocument.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${contentTypeOverrides.join('')}
</Types>`,
    });

    zipEntries.push({
      name: 'xl/workbook.xml',
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>${workbookSheetsXml.join('')}</sheets>
</workbook>`,
    });

    zipEntries.push({
      name: 'xl/_rels/workbook.xml.rels',
      content: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${workbookRelsXml.join('')}
</Relationships>`,
    });

    return createZipBlob(zipEntries);
  }

  function triggerDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 500);
  }

  const api = {
    readXlsx,
    readCsvBuffer,
    parseSpreadsheetFile,
    createXlsxBlob,
    triggerDownload,
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  } else {
    global.XlsxLite = api;
  }
})(typeof window !== 'undefined' ? window : globalThis);
