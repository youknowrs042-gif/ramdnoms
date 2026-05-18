#!/usr/bin/env python3
"""
Build a .docx file for qualitative coding results:
Manajemen Konflik Kepala Sekolah dan Teacher Well-Being
di SD Muhammadiyah Bodon
"""

import zipfile
import os
import sys

sys.path.insert(0, '/projects/sandbox/ramdnoms')
from coding_data_2 import coding_data

# ============================================================
# DOCX XML HELPERS
# ============================================================

def make_content_types():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

def make_rels():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

def make_word_rels():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

def make_styles():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:spacing w:before="360" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/><w:color w:val="1F3864"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:pPr><w:spacing w:before="240" w:after="100"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/><w:color w:val="2E75B6"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:pPr><w:spacing w:before="160" w:after="80"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr>
  </w:style>
</w:styles>'''

def escape_xml(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

def make_heading(text, level=1):
    style = f"Heading{level}"
    return f'<w:p><w:pPr><w:pStyle w:val="{style}"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>{escape_xml(text)}</w:t></w:r></w:p>'

def make_paragraph(text, bold=False, italic=False, highlight=None, size=22, color=None):
    rpr_parts = []
    if bold:
        rpr_parts.append("<w:b/>")
    if italic:
        rpr_parts.append("<w:i/>")
    if highlight:
        rpr_parts.append(f'<w:highlight w:val="{highlight}"/>')
    if color:
        rpr_parts.append(f'<w:color w:val="{color}"/>')
    rpr_parts.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
    rpr = f"<w:rPr>{''.join(rpr_parts)}</w:rPr>"
    return f'<w:p><w:r>{rpr}<w:t xml:space="preserve">{escape_xml(text)}</w:t></w:r></w:p>'

def make_mixed_paragraph(parts):
    runs = ""
    for part in parts:
        text = part.get("text", "")
        bold = part.get("bold", False)
        italic = part.get("italic", False)
        highlight = part.get("highlight", None)
        size = part.get("size", 22)
        color = part.get("color", None)
        rpr_parts = []
        if bold:
            rpr_parts.append("<w:b/>")
        if italic:
            rpr_parts.append("<w:i/>")
        if highlight:
            rpr_parts.append(f'<w:highlight w:val="{highlight}"/>')
        if color:
            rpr_parts.append(f'<w:color w:val="{color}"/>')
        rpr_parts.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
        rpr = f"<w:rPr>{''.join(rpr_parts)}</w:rPr>"
        runs += f'<w:r>{rpr}<w:t xml:space="preserve">{escape_xml(text)}</w:t></w:r>'
    return f"<w:p>{runs}</w:p>"

def make_empty():
    return "<w:p/>"

def table_start(col_widths):
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in col_widths)
    return f'''<w:tbl>
<w:tblPr>
<w:tblStyle w:val="TableGrid"/>
<w:tblW w:w="0" w:type="auto"/>
<w:tblBorders>
<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
</w:tblBorders>
</w:tblPr>
<w:tblGrid>{grid}</w:tblGrid>'''

def table_end():
    return "</w:tbl>"

def table_row(cells, header=False, highlights=None, shading=None):
    row = "<w:tr>"
    if header:
        row += "<w:trPr><w:tblHeader/></w:trPr>"
    for i, cell_text in enumerate(cells):
        hl = highlights[i] if highlights and i < len(highlights) else None
        sh = shading if header else None
        
        rpr_parts = []
        if header:
            rpr_parts.append("<w:b/>")
            rpr_parts.append('<w:color w:val="FFFFFF"/>')
        if hl:
            rpr_parts.append(f'<w:highlight w:val="{hl}"/>')
        rpr_parts.append('<w:sz w:val="20"/><w:szCs w:val="20"/>')
        rpr = f"<w:rPr>{''.join(rpr_parts)}</w:rPr>"
        
        tc_pr = ""
        if header:
            tc_pr = '<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="2E75B6"/></w:tcPr>'
        
        row += f'<w:tc>{tc_pr}<w:p><w:r>{rpr}<w:t xml:space="preserve">{escape_xml(cell_text)}</w:t></w:r></w:p></w:tc>'
    row += "</w:tr>"
    return row


# ============================================================
# BUILD DOCUMENT
# ============================================================

def build_document():
    body = ""
    
    # ===== HALAMAN JUDUL =====
    body += make_heading("HASIL KODING WAWANCARA KUALITATIF", 1)
    body += make_empty()
    body += make_paragraph("Judul Penelitian:", bold=True, size=24)
    body += make_paragraph("Manajemen Konflik Kepala Sekolah dan Teacher Well-Being di SD Muhammadiyah Bodon", size=24)
    body += make_empty()
    body += make_paragraph("Metode Analisis:", bold=True)
    body += make_paragraph("Open Coding \u2192 Axial Coding \u2192 Selective Coding (Tema Utama)")
    body += make_empty()
    body += make_paragraph("Sumber Data:", bold=True)
    body += make_paragraph("1. Wawancara Kepala Sekolah (KS) - Kepala SD Muhammadiyah Bodon")
    body += make_paragraph("2. Wawancara Guru (G) - Guru SD Muhammadiyah Bodon")
    body += make_empty()
    
    # Keterangan Warna
    body += make_paragraph("Keterangan Warna Koding (Highlight/Stabilo):", bold=True)
    body += make_mixed_paragraph([
        {"text": "\u25a0 KUNING (Yellow)", "highlight": "yellow", "bold": True},
        {"text": " = Tema 1: Profil dan Gaya Kepemimpinan Kepala Sekolah"}
    ])
    body += make_mixed_paragraph([
        {"text": "\u25a0 CYAN (Biru Muda)", "highlight": "cyan", "bold": True},
        {"text": " = Tema 2: Manajemen Konflik Kepala Sekolah"}
    ])
    body += make_mixed_paragraph([
        {"text": "\u25a0 HIJAU (Green)", "highlight": "green", "bold": True},
        {"text": " = Tema 3: Kondisi Teacher Well-Being"}
    ])
    body += make_mixed_paragraph([
        {"text": "\u25a0 MAGENTA (Ungu)", "highlight": "magenta", "bold": True},
        {"text": " = Tema 4: Hubungan Manajemen Konflik dan Teacher Well-Being"}
    ])
    body += make_empty()
    
    # ===== FOKUS PENELITIAN =====
    body += make_heading("FOKUS PENELITIAN", 1)
    body += make_paragraph("1. Mendeskripsikan profil dan gaya kepemimpinan kepala sekolah di SD Muhammadiyah Bodon.")
    body += make_paragraph("2. Menganalisis strategi manajemen konflik yang diterapkan kepala sekolah di SD Muhammadiyah Bodon.")
    body += make_paragraph("3. Menganalisis kondisi teacher well-being guru di SD Muhammadiyah Bodon.")
    body += make_paragraph("4. Menganalisis pengaruh manajemen konflik kepala sekolah terhadap teacher well-being guru.")
    body += make_empty()
    
    # ===== RINGKASAN STRUKTUR KODING =====
    body += make_heading("RINGKASAN STRUKTUR KODING", 1)
    body += make_empty()
    
    body += table_start([600, 3500, 4500, 1200])
    body += table_row(["No.", "Tema Utama (Selective Code)", "Axial Codes", "Jumlah Open Code"], header=True)
    body += table_row(["1", "Profil dan Gaya Kepemimpinan Kepala Sekolah",
                       "a. Filosofi dan Orientasi Kepemimpinan\nb. Gaya Kepemimpinan Demokratis\nc. Kepedulian dan Dukungan terhadap Guru\nd. Supervisi dan Pembinaan Profesional",
                       "17"], highlights=[None, "yellow", None, None])
    body += table_row(["2", "Manajemen Konflik Kepala Sekolah",
                       "a. Sumber dan Bentuk Konflik\nb. Deteksi dan Respons Awal terhadap Konflik\nc. Strategi Mediasi dan Penyelesaian Konflik\nd. Tindak Lanjut dan Keberlanjutan Resolusi",
                       "21"], highlights=[None, "cyan", None, None])
    body += table_row(["3", "Kondisi Teacher Well-Being",
                       "a. Dimensi Kesejahteraan Guru\nb. Faktor Pendukung Well-Being\nc. Faktor Penghambat dan Stressor\nd. Gejala Burnout dan Mekanisme Coping",
                       "15"], highlights=[None, "green", None, None])
    body += table_row(["4", "Hubungan Manajemen Konflik dan Teacher Well-Being",
                       "a. Dampak Konflik terhadap Well-Being Guru\nb. Peran Manajemen Konflik dalam Mendukung Well-Being\nc. Rekomendasi Perbaikan",
                       "14"], highlights=[None, "magenta", None, None])
    body += table_end()
    body += make_empty()
    body += make_paragraph(f"Total: 67 Open Codes | 15 Axial Codes | 4 Tema Utama", bold=True)
    body += make_empty()
    
    # ===== DETAIL PER TEMA =====
    themes_order = [
        "Profil dan Gaya Kepemimpinan Kepala Sekolah",
        "Manajemen Konflik Kepala Sekolah",
        "Kondisi Teacher Well-Being",
        "Hubungan Manajemen Konflik dan Teacher Well-Being"
    ]
    
    theme_objectives = {
        "Profil dan Gaya Kepemimpinan Kepala Sekolah": "Fokus 1: Mendeskripsikan profil dan gaya kepemimpinan kepala sekolah di SD Muhammadiyah Bodon.",
        "Manajemen Konflik Kepala Sekolah": "Fokus 2: Menganalisis strategi manajemen konflik yang diterapkan kepala sekolah.",
        "Kondisi Teacher Well-Being": "Fokus 3: Menganalisis kondisi teacher well-being guru di SD Muhammadiyah Bodon.",
        "Hubungan Manajemen Konflik dan Teacher Well-Being": "Fokus 4: Menganalisis pengaruh manajemen konflik kepala sekolah terhadap teacher well-being guru."
    }
    
    theme_colors = {
        "Profil dan Gaya Kepemimpinan Kepala Sekolah": "yellow",
        "Manajemen Konflik Kepala Sekolah": "cyan",
        "Kondisi Teacher Well-Being": "green",
        "Hubungan Manajemen Konflik dan Teacher Well-Being": "magenta"
    }
    
    for t_idx, theme in enumerate(themes_order, 1):
        color = theme_colors[theme]
        objective = theme_objectives[theme]
        
        body += make_heading(f"TEMA {t_idx}: {theme.upper()}", 1)
        body += make_paragraph(f"[{objective}]", italic=True)
        body += make_empty()
        
        theme_data = [d for d in coding_data if d["theme"] == theme]
        
        # Group by axial code preserving order
        axial_order = []
        axial_groups = {}
        for d in theme_data:
            ac = d["axial_code"]
            if ac not in axial_groups:
                axial_order.append(ac)
                axial_groups[ac] = []
            axial_groups[ac].append(d)
        
        for ac_idx, axial_code in enumerate(axial_order, 1):
            items = axial_groups[axial_code]
            
            body += make_heading(f"{t_idx}.{ac_idx} Axial Code: {axial_code}", 2)
            body += make_empty()
            
            # Table
            body += table_start([500, 2000, 4000, 1800, 1200])
            body += table_row(["No.", "Open Code", "Quotation (Kutipan Langsung)", "Sumber Data", "Warna"], header=True)
            
            for item_idx, item in enumerate(items, 1):
                body += table_row(
                    [
                        str(item_idx),
                        item["open_code"],
                        f'\u201c{item["quotation"]}\u201d',
                        item["source"],
                        color.upper()
                    ],
                    highlights=[None, color, None, None, color]
                )
            
            body += table_end()
            body += make_empty()
    
    # ===== MATRIKS KODING TERINTEGRASI =====
    body += make_heading("MATRIKS KODING TERINTEGRASI", 1)
    body += make_paragraph("Tabel berikut menampilkan seluruh hasil koding secara terintegrasi dari Open Code, Axial Code, hingga Tema Utama beserta kutipan dan sumber data.")
    body += make_empty()
    
    body += table_start([400, 1800, 1800, 3200, 1600, 700])
    body += table_row(["No.", "Tema Utama", "Axial Code", "Quotation", "Open Code", "Sumber"], header=True)
    
    for idx, item in enumerate(coding_data, 1):
        color = item["color"]
        # Shorten theme name for table
        theme_short = item["theme"]
        if len(theme_short) > 30:
            theme_short = theme_short[:28] + "..."
        
        # Shorten quotation for overview table
        quot = item["quotation"]
        if len(quot) > 100:
            quot = quot[:97] + "..."
        
        source_short = item["source"].replace("Wawancara Kepala Sekolah", "KS").replace("Wawancara Guru", "G")
        
        body += table_row(
            [str(idx), theme_short, item["axial_code"], f'\u201c{quot}\u201d', item["open_code"], source_short],
            highlights=[None, color, None, None, color, None]
        )
    
    body += table_end()
    body += make_empty()
    
    # ===== KESIMPULAN TEMUAN =====
    body += make_heading("KESIMPULAN TEMUAN", 1)
    body += make_empty()
    
    body += make_heading("Tema 1: Profil dan Gaya Kepemimpinan Kepala Sekolah", 2)
    body += make_paragraph("Kepala SD Muhammadiyah Bodon menerapkan gaya kepemimpinan demokratis yang berlandaskan nilai-nilai Islam berkemajuan Muhammadiyah (amanah, musyawarah, fastabiqul khairat). Kepemimpinan dimaknai sebagai pelayanan (servant leadership) dengan orientasi pada kehadiran aktif, keterbukaan, dan aksesibilitas. Dalam praktiknya, kepala sekolah melibatkan guru secara substantif dalam pengambilan keputusan, memberikan supervisi yang non-intimidatif dan reflektif, serta menunjukkan kepedulian personal terhadap kondisi emosional guru. Apresiasi diberikan secara langsung dan tulus sebagai bentuk motivasi non-material.")
    body += make_empty()
    
    body += make_heading("Tema 2: Manajemen Konflik Kepala Sekolah", 2)
    body += make_paragraph("Konflik yang terjadi di sekolah umumnya bersumber dari miskomunikasi, ketidakmerataan beban kerja, perbedaan karakter, dan erosi kepercayaan jangka panjang. Kepala sekolah mendeteksi konflik melalui perubahan suasana dan jaringan informasi guru senior, kemudian merespons dengan cepat tanpa menghakimi. Strategi penyelesaian mengutamakan pendekatan kolaboratif: mendengarkan semua pihak secara terpisah, menciptakan ruang aman untuk dialog, dan memfasilitasi pertemuan bersama sebagai mediator netral. Tantangan yang masih dihadapi adalah belum adanya SOP tertulis dan follow-up pasca-resolusi yang belum konsisten.")
    body += make_empty()
    
    body += make_heading("Tema 3: Kondisi Teacher Well-Being", 2)
    body += make_paragraph("Teacher well-being dipahami secara holistik mencakup dimensi emosional, sosial, profesional, dan fisik. Faktor pendukung utama meliputi kepuasan dari keberhasilan siswa, dukungan sosial rekan sejawat, kegiatan kebersamaan informal, dan kebijakan pintu terbuka kepala sekolah. Stressor utama adalah beban administrasi, tuntutan emosional menghadapi keberagaman siswa, kelelahan fisik saat periode intensif, dan keterbatasan sumber daya sekolah. Guru mengalami gejala burnout terutama menjelang akhir semester dengan strategi coping berupa pengambilan jeda, penetapan batasan kerja-rumah, dan berbagi dengan rekan terpercaya.")
    body += make_empty()
    
    body += make_heading("Tema 4: Hubungan Manajemen Konflik dan Teacher Well-Being", 2)
    body += make_paragraph("Terdapat hubungan kuat antara manajemen konflik dan teacher well-being. Konflik yang tidak terselesaikan berdampak langsung pada penurunan fokus mengajar, kreativitas, semangat kerja, dan bahkan kesehatan fisik guru. Sebaliknya, resolusi konflik yang cepat dan adil oleh kepala sekolah mampu memulihkan suasana kondusif dan meningkatkan produktivitas. Rekomendasi perbaikan mencakup: penyusunan SOP konflik yang tersosialisasi, follow-up aktif pasca-resolusi, pembentukan forum ekspresi rutin untuk pencegahan, konsistensi kepemimpinan untuk membangun kepercayaan, serta solusi berbasis non-material seperti komunikasi terbuka dan budaya saling mendukung.")
    body += make_empty()
    
    # Build full XML document
    document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
{body}
<w:sectPr>
<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
</w:sectPr>
</w:body>
</w:document>'''
    
    return document_xml


def create_docx(output_path):
    print("Building document content...")
    document_xml = build_document()
    
    print("Writing .docx file...")
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', make_content_types())
        zf.writestr('_rels/.rels', make_rels())
        zf.writestr('word/_rels/document.xml.rels', make_word_rels())
        zf.writestr('word/document.xml', document_xml)
        zf.writestr('word/styles.xml', make_styles())
    
    size = os.path.getsize(output_path)
    print(f"\nDocument created successfully!")
    print(f"  File: {output_path}")
    print(f"  Size: {size:,} bytes")


if __name__ == "__main__":
    output = "/projects/sandbox/ramdnoms/Hasil_Koding_Manajemen_Konflik_Well_Being.docx"
    create_docx(output)
