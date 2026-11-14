"""
Sıfır dış bağımlılıkla (sadece Python standart kütüphanesi ile)
gerçekçi inşaat muhasebesi örnek Excel (.xlsx) dosyaları üretir.
"""
import os
import zipfile
from xml.sax.saxutils import escape


def col_letter(idx: int) -> str:
    n = idx + 1
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def write_simple_xlsx(path: str, sheet_name: str, headers: list[str], rows: list[list]):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    all_rows = [headers] + rows

    rows_xml = []
    for r_idx, row in enumerate(all_rows, start=1):
        cells_xml = []
        for c_idx, val in enumerate(row):
            ref = f"{col_letter(c_idx)}{r_idx}"
            style = "1" if r_idx == 1 else ("2" if isinstance(val, (int, float)) else "0")
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                cells_xml.append(f'<c r="{ref}" s="{style}"><v>{val}</v></c>')
            else:
                cells_xml.append(
                    f'<c r="{ref}" t="inlineStr" s="{style}"><is><t>{escape(str(val))}</t></is></c>'
                )
        rows_xml.append(f'<row r="{r_idx}">{"".join(cells_xml)}</row>')

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(rows_xml)}</sheetData>'
        '</worksheet>'
    )

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts>'
        '<fonts count="2">'
        '<font><sz val="10"/><color rgb="FF0F172A"/><name val="Calibri"/></font>'
        '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
        '</fonts>'
        '<fills count="3">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF1E293B"/></patternFill></fill>'
        '</fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="3">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" applyFont="1" applyFill="1"/>'
        '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/>'
        '</cellXfs>'
        '</styleSheet>'
    )

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-officedocument.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>'
    )

    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )

    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets><sheet name="{escape(sheet_name)}" sheetId="1" r:id="rId1"/></sheets>'
        '</workbook>'
    )

    wb_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '</Relationships>'
    )

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        zf.writestr("xl/styles.xml", styles_xml)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def main():
    base_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ornek-dosyalar")

    irs_headers = [
        "İrsaliye No", "Tarih", "Şantiye", "Tedarikçi",
        "Malzeme Cinsi", "Birim", "Miktar", "Birim Fiyat", "Toplam Tutar"
    ]
    irs_rows = [
        ["IRS-2026-101", "02.11.2026", "Güneşli Konutları", "Anadolu Hazır Beton A.Ş.", "C30 Hazır Beton", "m³", 40, 2850, 114000],
        ["IRS-2026-102", "03.11.2026", "Güneşli Konutları", "Anadolu Hazır Beton A.Ş.", "C30 Hazır Beton", "m³", 24, 2850, 68400],
        ["IRS-2026-103", "04.11.2026", "Kartal Rezidans", "DemirÇelik Yapı San. Tic.", "Ø14 Nervürlü İnşaat Demiri", "Ton", 15, 24200, 363000],
        ["IRS-2026-104", "05.11.2026", "Güneşli Konutları", "Anadolu Hazır Beton A.Ş.", "C30 Hazır Beton", "m³", 30, 2850, 85500],
        ["IRS-2026-105", "06.11.2026", "Ataşehir İş Merkezi", "Kaya Agrega & Yapı Malz.", "Kırma Taş Mıcır (15-25 mm)", "Ton", 85, 420, 35700],
        ["IRS-2026-106", "07.11.2026", "Kartal Rezidans", "Marmara Nakliyat Ltd.", "Hafriyat & Moloz Nakliyesi", "Sefer", 12, 3200, 38400],
        ["IRS-2026-107", "08.11.2026", "Ataşehir İş Merkezi", "Anadolu Hazır Beton A.Ş.", "C35 Hazır Beton + Pompa", "m³", 52, 3150, 163800],
        ["IRS-2026-108", "09.11.2026", "Kartal Rezidans", "DemirÇelik Yapı San. Tic.", "Ø16 Nervürlü İnşaat Demiri", "Ton", 18, 24200, 435600],
        ["IRS-2026-109", "10.11.2026", "Güneşli Konutları", "Kaya Agrega & Yapı Malz.", "Yıkanmış Sıva Kumu", "Ton", 60, 480, 28800],
        ["IRS-2026-110", "11.11.2026", "Ataşehir İş Merkezi", "DemirÇelik Yapı San. Tic.", "Q188 Çelik Hasır", "Adet", 140, 690, 96600],
        ["IRS-2026-111", "11.11.2026", "Güneşli Konutları", "Marmara Nakliyat Ltd.", "Malzeme Vinç & Nakliye", "Saat", 16, 1950, 31200],
        ["IRS-2026-112", "12.11.2026", "Kartal Rezidans", "Anadolu Hazır Beton A.Ş.", "C30 Hazır Beton", "m³", 45, 2850, 128250],
        ["IRS-2026-113", "13.11.2026", "Ataşehir İş Merkezi", "Kaya Agrega & Yapı Malz.", "Bims Blok 20cm", "Adet", 2500, 22, 55000],
        ["IRS-2026-114", "14.11.2026", "Ataşehir İş Merkezi", "Kaya Agrega & Yapı Malz.", "CEM I 42.5 Çimento", "Torba", 160, 175, 28000],
        ["IRS-2026-115", "14.11.2026", "Güneşli Konutları", "DemirÇelik Yapı San. Tic.", "Ø12 Nervürlü İnşaat Demiri", "Ton", 6, 24200, 145200],
    ]

    fat_headers = [
        "Fatura No", "İrsaliye No", "Tarih", "Şantiye", "Tedarikçi",
        "Malzeme Cinsi", "Birim", "Miktar", "Birim Fiyat", "Toplam Tutar"
    ]
    fat_rows = [
        ["FAT-4410", "IRS-2026-101", "05.11.2026", "Güneşli Konutları", "Anadolu Hazır Beton A.Ş.", "C30 Hazır Beton", "m³", 40, 2850, 114000],
        ["FAT-4412", "IRS-2026-102", "05.11.2026", "Güneşli Konutları", "Anadolu Hazır Beton A.Ş.", "C30 Hazır Beton", "m³", 24, 2850, 68400],
        ["FAT-7021", "IRS-2026-103", "06.11.2026", "Kartal Rezidans", "DemirÇelik Yapı San. Tic.", "Ø14 Nervürlü İnşaat Demiri", "Ton", 15, 24200, 363000],
        ["FAT-4418", "IRS-2026-104", "07.11.2026", "Güneşli Konutları", "Anadolu Hazır Beton A.Ş.", "C30 Hazır Beton", "m³", 36, 2850, 102600],
        ["FAT-4419", "IRS-2026-102", "08.11.2026", "Güneşli Konutları", "Anadolu Hazır Beton A.Ş.", "C30 Hazır Beton", "m³", 24, 2850, 68400],
        ["FAT-3190", "IRS-2026-105", "08.11.2026", "Ataşehir İş Merkezi", "Kaya Agrega & Yapı Malz.", "Kırma Taş Mıcır (15-25 mm)", "Ton", 85, 420, 35700],
        ["FAT-8830", "IRS-2026-106", "09.11.2026", "Kartal Rezidans", "Marmara Nakliyat Ltd.", "Hafriyat & Moloz Nakliyesi", "Sefer", 12, 3200, 38400],
        ["FAT-4425", "IRS-2026-107", "10.11.2026", "Ataşehir İş Merkezi", "Anadolu Hazır Beton A.Ş.", "C35 Hazır Beton + Pompa", "m³", 52, 3150, 163800],
        ["FAT-7033", "IRS-2026-108", "11.11.2026", "Kartal Rezidans", "DemirÇelik Yapı San. Tic.", "Ø16 Nervürlü İnşaat Demiri", "Ton", 18, 25400, 457200],
        ["FAT-3204", "IRS-2026-109", "11.11.2026", "Güneşli Konutları", "Kaya Agrega & Yapı Malz.", "Yıkanmış Sıva Kumu", "Ton", 60, 480, 28800],
        ["FAT-7040", "IRS-2026-110", "12.11.2026", "Ataşehir İş Merkezi", "DemirÇelik Yapı San. Tic.", "Q188 Çelik Hasır", "Adet", 140, 690, 96600],
        ["FAT-8832", "IRS-2026-111", "12.11.2026", "Güneşli Konutları", "Marmara Nakliyat Ltd.", "Malzeme Vinç & Nakliye", "Saat", 16, 1950, 31200],
        ["FAT-4439", "IRS-2026-112", "13.11.2026", "Kartal Rezidans", "Anadolu Hazır Beton A.Ş.", "C30 Hazır Beton", "m³", 45, 2850, 128250],
        ["FAT-3211", "IRS-2026-113", "14.11.2026", "Ataşehir İş Merkezi", "Kaya Agrega & Yapı Malz.", "Bims Blok 20cm", "Adet", 2500, 22, 55000],
        ["FAT-8834", "IRS-2026-999", "14.11.2026", "Kartal Rezidans", "Marmara Nakliyat Ltd.", "Ekstra Kamyon Seferi (Onaysız)", "Sefer", 5, 3700, 18500],
    ]

    write_simple_xlsx(
        os.path.join(base_dir, "santiye_irsaliyeleri.xlsx"),
        "İrsaliyeler",
        irs_headers,
        irs_rows,
    )
    write_simple_xlsx(
        os.path.join(base_dir, "tedarikci_faturalari.xlsx"),
        "Faturalar",
        fat_headers,
        fat_rows,
    )
    print("Örnek Excel dosyaları oluşturuldu:", base_dir)


if __name__ == "__main__":
    main()
