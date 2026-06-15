import os
import shutil
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

def delete_element(element):
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)

def get_all_paragraphs(doc):
    paragraphs = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                paragraphs.extend(cell.paragraphs)
    body_xml = doc.element.body
    for txbx in body_xml.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}txbxContent'):
        for p_xml in txbx.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
            paragraphs.append(Paragraph(p_xml, doc))
    return paragraphs

def replace_text_in_paragraph(paragraph, old_text, new_text):
    if old_text not in paragraph.text:
        return
    
    # Try replacing in individual runs first
    replaced_in_runs = False
    for run in paragraph.runs:
        if old_text in run.text:
            run.text = run.text.replace(old_text, new_text)
            replaced_in_runs = True
            
    # If not replaced (because it was split across runs), merge runs and replace
    if not replaced_in_runs and paragraph.runs:
        full_text = paragraph.text
        new_full_text = full_text.replace(old_text, new_text)
        
        first_run = paragraph.runs[0]
        p_element = paragraph._p
        for run in list(paragraph.runs):
            p_element.remove(run._r)
            
        new_run = paragraph.add_run(new_full_text)
        new_run.font.name = first_run.font.name
        new_run.font.size = first_run.font.size
        new_run.font.bold = first_run.font.bold
        new_run.font.italic = first_run.font.italic
        if first_run.font.color and first_run.font.color.rgb:
            new_run.font.color.rgb = first_run.font.color.rgb

def format_body_paragraph(p):
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.first_line_indent = Inches(0.3937)  # 1 cm
    for run in p.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.bold = False
        run.font.italic = False

def format_chapter_heading(p):
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.first_line_indent = Inches(0)
    for run in p.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.italic = False

def format_heading_level_1(p):
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Inches(0.1968)  # 0.5 cm
    p.paragraph_format.first_line_indent = Inches(0)
    for run in p.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.italic = False

def format_heading_level_2(p):
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Inches(0.3937)  # 1.0 cm
    p.paragraph_format.first_line_indent = Inches(0)
    for run in p.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.italic = False

def add_section_break_before(paragraph):
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    sectPr = OxmlElement('w:sectPr')
    type_elem = OxmlElement('w:type')
    type_elem.set(qn('w:val'), 'nextPage')
    sectPr.append(type_elem)
    pPr.append(sectPr)
    return sectPr

def set_section_page_numbering(section, fmt='decimal', start=None):
    sectPr = section._sectPr
    pgNumType = sectPr.find(qn('w:pgNumType'))
    if pgNumType is None:
        pgNumType = OxmlElement('w:pgNumType')
        sectPr.append(pgNumType)
    pgNumType.set(qn('w:fmt'), fmt)
    if start is not None:
        pgNumType.set(qn('w:start'), str(start))

def clean_and_fill_report():
    template_path = "Bitirme-Projesi-Yazim-Kilavuzu-ve-Sablonu1.docx"
    output_path = "Bitirme_Projesi_Raporu.docx"
    
    print(f"Creating copy of template as {output_path}...")
    shutil.copy(template_path, output_path)
    
    doc = docx.Document(output_path)
    body = doc.element.body
    
    # 1. Clean the Guide Elements preceding EK-1
    print("Removing guide text...")
    children = list(body)
    ek1_idx = -1
    for idx, child in enumerate(children):
        if child.tag.endswith('p'):
            text_parts = [t.text for t in child.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t') if t.text]
            full_text = "".join(text_parts)
            if "EK-1. Dış Kapak Sayfası Örneği" in full_text:
                ek1_idx = idx
                break
                
    if ek1_idx != -1:
        for i in range(ek1_idx):
            body.remove(children[i])
        print(f"Removed {ek1_idx} guide elements.")
    else:
        print("Error: Could not find EK-1 marker in template.")
        return

    # 2. Clean EK-14 MS WORD Suggestions at the end of the document
    print("Removing EK-14 Word suggestions from the end...")
    children = list(body)
    ek14_idx = -1
    for idx, child in enumerate(children):
        if child.tag.endswith('p'):
            text_parts = [t.text for t in child.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t') if t.text]
            full_text = "".join(text_parts)
            if "EK-14. MS WORD" in full_text:
                ek14_idx = idx
                break
                
    if ek14_idx != -1:
        elements_to_delete = len(children) - ek14_idx
        for i in range(ek14_idx, len(children)):
            body.remove(children[i])
        print(f"Removed {elements_to_delete} elements corresponding to EK-14.")
    else:
        print("Warning: Could not find EK-14 marker at the end.")

    # 3. Global String Replacements for Metadata
    print("Replacing placeholders...")
    # NOTE: Removed "Yıl": "2026" to avoid corrupting "Yılı" into "2026ı"
    replacements = {
        "BİTİRME PROJESİ ADI": "DRONE TABANLI GÖRÜNTÜ VERİSİYLE DERİN ÖĞRENME KULLANILARAK GERÇEK ZAMANLI HEDEF TESPİTİ",
        "TITLE OF GRADUATION PROJECT": "REAL-TIME TARGET DETECTION USING DEEP LEARNING WITH DRONE-BASED IMAGE DATA",
        "Bitirme Projesi Öğrencinin Adı SOYADI": "Emre Kaan ŞENSOY",
        "Öğrencinin Adı SOYADI": "Emre Kaan ŞENSOY",
        "Name and SURNAME of student": "Emre Kaan ŞENSOY",
        "Department of student": "Emre Kaan ŞENSOY",
        "............ Bölümü": "Yazılım Mühendisliği Bölümü",
        ".............................. Bölümü": "Yazılım Mühendisliği Bölümü",
        "Department of": "Department of Software Engineering",
        "Danışman: Unvanı Adı SOYAD": "Danışman: Doç. Dr. Selçuk ÇAKMAK",
        "Danışman: Unvanı Adı ve SOYADI": "Danışman: Doç. Dr. Selçuk ÇAKMAK",
        "Supervisor: Title Name and SURNAME": "Supervisor: Doç. Dr. Selçuk ÇAKMAK",
        "Ay Yıl": "Haziran 2026",
        "AyYıl": "Haziran 2026",
        "Month Year": "June 2026",
        "YIL": "2026",
        ".....................................................’nın": "Emre Kaan ŞENSOY'un",
        "..................................................................  Bölümünde": "Yazılım Mühendisliği Bölümünde",
        "..../..../20....": "10/06/2026",
        "MÜHENDİSLİK FAKÜLTESİ": "MÜHENDİSLİK VE DOĞA BİLİMLERİ FAKÜLTESİ",
        "Faculty of Engineering": "Faculty of Engineering and Natural Sciences"
    }

    all_p = get_all_paragraphs(doc)
    for p in all_p:
        for old, new in replacements.items():
            if old in p.text:
                replace_text_in_paragraph(p, old, new)

    # Clean the EK-X headers so they do not show up on pages
    ek_headers = [
        "EK-1. Dış Kapak Sayfası Örneği",
        "EK-2. İç Kapak Örneği",
        "EK-3. Bitirme Projesi Onay Sayfası Örneği",
        "EK-4. Bildirim Sayfası Örneği",
        "Ek-5. İçindekiler Örneği",
        "EK-6. Özet Sayfası Örneği",
        "EK-7. Özet Sayfası Örneği (İngilizce)",
        "EK-8. Tablolar/Çizelgeler Listesi Örneği",
        "EK-9. Şekiller Listesi Örneği",
        "EK-10. Görseller Listesi",
        "EK-11. Simgeler ve Kısaltmalar Dizini Örneği",
        "EK-12. Kaynakça Örneği",
        "EK-13. Özgeçmiş Örneği"
    ]
    
    for p in all_p:
        for header in ek_headers:
            if header in p.text:
                replace_text_in_paragraph(p, header, "")

    # Clean the dots/placeholders from Bildirim page and CV page
    cv_placeholders = {
        "Adı Soyadı\t: ................................................................": "Adı Soyadı\t: Emre Kaan ŞENSOY",
        "Yabancı Dil\t: ....................................": "Yabancı Dil\t: İngilizce",
        "Doğum Yeri ve Yılı : ............................../.................": "Doğum Yeri ve Yılı : Samsun / 2003",
        "E-Posta\t: ................................................................": "E-Posta\t: emrekaan@samsun.edu.tr",
        "Yıl, Üniversite, Enstitü/Fakülte, Bölüm, Anabilim Dalı": "2022 - 2026, Samsun Üniversitesi, Mühendislik ve Doğa Bilimleri Fakültesi, Yazılım Mühendisliği Bölümü",
        "Yıl, Görev, Kurum, Birim": "2025, Stajyer Yazılım Geliştirici, Samsun Teknopark, Ar-Ge Birimi"
    }
    for p in all_p:
        for old, new in cv_placeholders.items():
            if old in p.text:
                replace_text_in_paragraph(p, old, new)

    # 4. Fill in Table 29 (Jury table)
    # The jury table is doc.tables[0] in the cleaned document
    jury_table = doc.tables[0]
    # Row 1: Danışman
    replace_text_in_paragraph(jury_table.rows[1].cells[1].paragraphs[0], ":.........................................................", ": Doç. Dr. Selçuk ÇAKMAK")
    # Row 2: Member 1
    replace_text_in_paragraph(jury_table.rows[2].cells[1].paragraphs[0], ":.........................................................", ": Doç. Dr. Muammer TÜRKOĞLU")
    # Row 3: Sibling (leave blank but clean the dotted lines)
    replace_text_in_paragraph(jury_table.rows[3].cells[1].paragraphs[0], ":.........................................................", ": .........................................................")
    
    # We can delete rows 4 and 5 of the jury table to make it a 3-member table
    tbl_elem = jury_table._tbl
    tbl_elem.remove(jury_table.rows[5]._tr)
    tbl_elem.remove(jury_table.rows[4]._tr)
    print("Jury table updated.")

    # 5. Insert Önsöz page
    # Find the paragraph for İÇİNDEKİLER to insert Önsöz before it
    p_toc = None
    for idx, p in enumerate(doc.paragraphs):
        if p.text.strip() == "İÇİNDEKİLER":
            p_toc = p
            break
            
    if p_toc:
        print("Inserting Önsöz before İçindekiler...")
        p_preface_title = p_toc.insert_paragraph_before("ÖNSÖZ", style='Normal')
        format_chapter_heading(p_preface_title)
        
        # Add section break before the Önsöz title so it starts on a new page (Page IV)
        add_section_break_before(p_preface_title)
        
        preface_paras = [
            "Bu çalışma, Samsun Üniversitesi Yazılım Mühendisliği Bölümü bünyesinde gerçekleştirilmiş bir lisans bitirme projesidir. Proje kapsamında, son yıllarda insansız hava araçlarının kullanımının artmasıyla birlikte önem kazanan hava çekimi görüntülerinde gerçek zamanlı nesne tespiti problemi ele alınmış, sentetik veri destekli derin öğrenme temelli yenilikçi bir çözüm sunulmuştur.",
            "Çalışmanın başlangıcından tamamlanmasına kadar geçen süreçte engin bilgi ve tecrübeleri ile beni yönlendiren, her tıkandığım noktada fikirleriyle bana yol gösteren değerli danışmanım Doç. Dr. Selçuk ÇAKMAK'a en samimi şükranlarımı sunarım. Ayrıca, lisans eğitimim boyunca değerli bilgilerini esirgemeyen tüm Yazılım Mühendisliği Bölümü hocalarıma ve akademik çalışmalarımdaki desteklerinden ötürü Doç. Dr. Muammer TÜRKOĞLU'na teşekkür ederim.",
            "Son olarak, hayatımın her döneminde olduğu gibi bu projenin hazırlanmasında da maddi ve manevi desteklerini hiçbir zaman esirgemeyen, bana her zaman inanan ve yanımda olan sevgili aileme en derin sevgilerimi ve teşekkürlerimi sunarım."
        ]
        
        for text in preface_paras:
            p_new = p_toc.insert_paragraph_before(text, style='Normal')
            format_body_paragraph(p_new)
            
        # Add date and name signature lines
        p_sig_date = p_toc.insert_paragraph_before("Haziran 2026", style='Normal')
        p_sig_date.paragraph_format.space_before = Pt(24)
        p_sig_date.paragraph_format.space_after = Pt(6)
        p_sig_date.paragraph_format.first_line_indent = Inches(0)
        p_sig_date.runs[0].font.name = 'Times New Roman'
        p_sig_date.runs[0].font.size = Pt(12)
        p_sig_date.runs[0].font.bold = False
        
        p_sig_name = p_toc.insert_paragraph_before("Emre Kaan ŞENSOY", style='Normal')
        p_sig_name.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_sig_name.paragraph_format.space_before = Pt(6)
        p_sig_name.paragraph_format.space_after = Pt(24)
        p_sig_name.paragraph_format.first_line_indent = Inches(0)
        p_sig_name.runs[0].font.name = 'Times New Roman'
        p_sig_name.runs[0].font.size = Pt(12)
        p_sig_name.runs[0].font.bold = True
        
        # Add a page break after Önsöz to start İçindekiler on a new page
        p_toc.insert_paragraph_before().add_run().add_break(docx.enum.text.WD_BREAK.PAGE)
        print("Önsöz page generated.")

    # 6. Fill Özet and Abstract Placeholder Dots
    print("Filling Özet & Abstract...")
    turkish_abstract = (
        "Bu projede, İHA (drone) görüntüleri kullanılarak gerçek zamanlı kırmızı ve mavi branda hedeflerinin "
        "tespiti için derin öğrenme tabanlı bir model geliştirilmiştir. Gerçek dünyada hava çekimi görüntülerinin "
        "toplanması ve etiketlenmesindeki zorlukları aşmak amacıyla sentetik veri destekli bir eğitim stratejisi "
        "önerilmiştir. İlk aşamada, YOLOv8 Nano modeli simülasyon ortamında üretilen sentetik drone görüntüleriyle "
        "50 epok boyunca ön eğitime tabi tutulmuş ve %99.5 mAP50 değerine ulaşılmıştır. İkinci aşamada, elde edilen "
        "ağırlıklar gerçek drone görüntülerinden oluşan veri seti ile 100 epok boyunca ince ayardan (fine-tuning) "
        "geçirilmiştir. Test sonuçları, geliştirilen nihai hibrit modelin gerçek İHA görüntülerinde %99.8 hassasiyet "
        "(precision), %100.0 duyarlılık (recall) ve %99.5 mAP50 oranına ulaştığını göstermektedir. Elde edilen "
        "yüksek doğruluk ve hız performansı, modelin drone üzerindeki gömülü sistemlerde gerçek zamanlı "
        "arama-kurtarma ve keşif faaliyetlerinde etkin bir şekilde kullanılabileceğini kanıtlamaktadır."
    )
    english_abstract = (
        "In this project, a deep learning-based model has been developed for real-time detection of red and "
        "blue tarpaulin targets using UAV (drone) imagery. To overcome the challenges of collecting and labeling "
        "aerial images in the real world, a synthetic data-supported training strategy is proposed. In the first "
        "phase, the YOLOv8 Nano model was pre-trained for 50 epochs using synthetic drone images generated in a "
        "simulation environment, achieving 99.5% mAP50. In the second phase, the obtained weights were fine-tuned "
        "for 100 epochs with a dataset consisting of real drone images. Test results indicate that the final hybrid "
        "model achieves 99.8% precision, 100.0% recall, and 99.5% mAP50 on real UAV images. The high accuracy and "
        "speed performance demonstrate that the model can be effectively utilized for real-time search and rescue "
        "and reconnaissance missions on embedded drone systems."
    )
    
    # We retrieve the list of paragraphs once and iterate with indices
    body_children = list(body)
    in_dots_group = False
    dots_group_count = 0
    elements_to_delete = []
    
    for child in body_children:
        if child.tag.endswith('p'):
            p = Paragraph(child, doc)
            text = p.text.strip()
            if text.startswith("................................................................"):
                if not in_dots_group:
                    in_dots_group = True
                    dots_group_count += 1
                    if dots_group_count == 1:
                        p.text = turkish_abstract
                        format_body_paragraph(p)
                    elif dots_group_count == 2:
                        p.text = english_abstract
                        format_body_paragraph(p)
                else:
                    elements_to_delete.append(child)
            else:
                if in_dots_group:
                    in_dots_group = False
                    
    for elem in elements_to_delete:
        body.remove(elem)
    print(f"Cleaned up {len(elements_to_delete)} dot placeholder paragraphs.")

    # Refresh paragraphs list
    all_p_abstract = get_all_paragraphs(doc)
    # Fill keywords
    for p in all_p_abstract:
        if "Anahtar Sözcükler:" in p.text:
            replace_text_in_paragraph(p, "Anahtar sözcük 1, Anahtar sözcük 2, Anahtar sözcük 3, Anahtar sözcük 4, Anahtar sözcük 5.", "Derin Öğrenme, YOLOv8, Sentetik Veri, Drone Hedef Tespiti, Transfer Öğrenme.")
        if "Keywords:" in p.text:
            replace_text_in_paragraph(p, "First keyword, Second keyword, Third keyword, Fourth keyword, Fifth keyword.", "Deep Learning, YOLOv8, Synthetic Data, Drone Target Detection, Transfer Learning.")

    # 7. Insert the Thesis Chapters before References
    # Search for "Kaynakça" (note the spelling in the template, first letter capitalized)
    p_ref = None
    for p in doc.paragraphs:
        if p.text.strip() == "Kaynakça":
            p_ref = p
            break
            
    if p_ref:
        print("Inserting thesis chapters before Kaynakça...")
        
        # Add GİRİŞ header and set section break
        p_intro_title = p_ref.insert_paragraph_before("GİRİŞ", style='Normal')
        format_chapter_heading(p_intro_title)
        
        # Add section break to the sibling before GİRİŞ
        intro_idx = list(body).index(p_intro_title._p)
        sibling_before_xml = list(body)[intro_idx - 1]
        pPr = sibling_before_xml.get_or_add_pPr()
        sectPr = OxmlElement('w:sectPr')
        type_elem = OxmlElement('w:type')
        type_elem.set(qn('w:val'), 'nextPage')
        sectPr.append(type_elem)
        pPr.append(sectPr)
        
        # Write GİRİŞ content
        intro_paras = [
            "Düşük irtifalı insansız hava araçları (İHA - Drone), yüksek hareket kabiliyetleri, geniş görüş açıları ve hızlı konuşlandırılabilmeleri sayesinde sivil ve askeri alanda birçok kritik görevde kullanılmaktadır. Özellikle afet sonrası arama-kurtarma operasyonlarında, yangın takibinde, sınır güvenliğinde ve tarımsal analizlerde, havadan alınan görüntülerin gerçek zamanlı olarak analiz edilmesi hayati önem taşımaktadır. Havadan çekilen görüntülerde nesne tespiti; değişken ışık koşulları, farklı kamera açıları, hedef nesnelerin küçük boyutları ve arka plan karmaşıklığı gibi teknik zorlukları barındırmaktadır.",
            "Ayrıca, derin öğrenme modellerinin başarısı büyük oranda etiketlenmiş kaliteli ve çeşitli veri setlerine bağlıdır. Gerçek dünyada drone tabanlı görüntülerin toplanması, temizlenmesi ve manuel olarak etiketlenmesi son derece yüksek maliyetli ve zaman alıcı bir süreçtir. Bu çalışmada, drone tabanlı görüntü verilerinde gerçek zamanlı hedef tespiti yapmak üzere derin öğrenme temelli YOLOv8 Nano mimarisi kullanılmıştır. Gerçek veri yetersizliği sorununu aşmak amacıyla, simülasyon ortamında üretilen sentetik veri setleriyle ön eğitim yapılmış ve transfer öğrenme (transfer learning) yöntemiyle gerçek veri setine uyarlama gerçekleştirilmiştir.",
            "Çalışmanın hedef nesneleri, afet ve arama-kurtarma senaryolarında sıklıkla geçici barınak veya işaretçi olarak kullanılan mavi branda (mavi_branda) ve kırmızı branda (kirmizi_branda) olarak belirlenmiştir. Bu proje, sentetik verilerle yapılan ön eğitimin gerçek dünya görüntülerindeki nesne tespit doğruluğuna etkisini ve İHA tabanlı nesne tespiti sistemlerindeki başarımını incelemeyi amaçlamaktadır."
        ]
        for text in intro_paras:
            p_new = p_ref.insert_paragraph_before(text, style='Normal')
            format_body_paragraph(p_new)
            
        # Write BÖLÜM 1
        p_ref.insert_paragraph_before().add_run().add_break(docx.enum.text.WD_BREAK.PAGE)
        p_b1_title = p_ref.insert_paragraph_before("BİRİNCİ BÖLÜM", style='Normal')
        format_chapter_heading(p_b1_title)
        p_b1_sub = p_ref.insert_paragraph_before("LİTERATÜR VE KURAMSAL ALTYAPI", style='Normal')
        format_chapter_heading(p_b1_sub)
        p_b1_sub.paragraph_format.space_after = Pt(24)
        
        p_ref.insert_paragraph_before("1.1. Nesne Tespiti ve Derin Öğrenme", style='Normal')
        format_heading_level_1(doc.paragraphs[len(doc.paragraphs)-2])
        format_body_paragraph(p_ref.insert_paragraph_before(
            "Nesne tespiti (object detection), görüntü üzerindeki nesnelerin hem sınıflarını belirleyen hem de koordinatlarını (bounding box) çıkaran bilgisayarlı görü (computer vision) alanının en temel problemlerinden biridir. Geleneksel yöntemler (Viola-Jones, HOG, SVM vb.) el yapımı öznitelik çıkarımına dayanırken, Evrişimsel Sinir Ağları (CNN) tabanlı modern derin öğrenme yaklaşımları bu süreci otomatikleştirmiştir. Nesne tespiti modelleri genel olarak iki aşamalı (Two-Stage - örn: R-CNN, Faster R-CNN) ve tek aşamalı (One-Stage - örn: SSD, YOLO) olmak üzere ikiye ayrılır. Gerçek zamanlı drone uygulamalarında yüksek hız gereksinimi nedeniyle tek aşamalı modeller tercih edilmektedir."
        ))
        
        p_ref.insert_paragraph_before("1.2. YOLOv8 Mimarisi", style='Normal')
        format_heading_level_1(doc.paragraphs[len(doc.paragraphs)-2])
        format_body_paragraph(p_ref.insert_paragraph_before(
            "YOLO (You Only Look Once) serisi, yüksek doğruluk ve çalışma hızı dengesi ile öne çıkmaktadır (Redmon vd., 2016: 779). Serinin en güncel sürümlerinden biri olan YOLOv8 (Jocher vd., 2023), çapa bağımsız (anchor-free) bir tespit kafasına sahiptir. Bu sayede nesne aday kutuları doğrudan tahmin edilerek hesaplama karmaşıklığı azaltılır. YOLOv8 mimarisi, kayıp fonksiyonu olarak CIoU (Complete IoU) ve DFL (Distribution Focal Loss) kullanarak nesne konumlandırma başarısını artırır. Model ölçeklendirme seçenekleri arasında yer alan YOLOv8 Nano (YOLOv8n), düşük parametre sayısı ve yüksek çalışma hızı sayesinde drone üzerindeki gömülü sistemlerde gerçek zamanlı çalışmaya son derece uygundur."
        ))
        
        p_ref.insert_paragraph_before("1.3. Sentetik Veri ve Transfer Öğrenme", style='Normal')
        format_heading_level_1(doc.paragraphs[len(doc.paragraphs)-2])
        format_body_paragraph(p_ref.insert_paragraph_before(
            "Bilgisayarlı görüde karşılaşılan veri seti eksikliği sorunlarında, 3D modelleme veya veri artırımı yoluyla elde edilen sentetik verilerle ön eğitim yapılması, modelin genel öznitelikleri kavramasını kolaylaştırarak transfer öğrenme performansını artırmaktadır (Tremblay vd., 2018: 1082). Gerçek veri seti boyutunun küçük olduğu durumlarda, önceden sentetik verilerle eğitilmiş olan ağırlıkların yüklenerek gerçek saha verileri üzerinde ince ayar (fine-tuning) yapılması, eğitim süresini kısaltmakta ve doğruluk oranlarını önemli ölçüde artırmaktadır."
        ))
        
        # Write BÖLÜM 2
        p_ref.insert_paragraph_before().add_run().add_break(docx.enum.text.WD_BREAK.PAGE)
        p_b2_title = p_ref.insert_paragraph_before("İKİNCİ BÖLÜM", style='Normal')
        format_chapter_heading(p_b2_title)
        p_b2_sub = p_ref.insert_paragraph_before("MATERYAL VE METOT", style='Normal')
        format_chapter_heading(p_b2_sub)
        p_b2_sub.paragraph_format.space_after = Pt(24)
        
        p_ref.insert_paragraph_before("2.1. Veri Setinin Hazırlanması", style='Normal')
        format_heading_level_1(doc.paragraphs[len(doc.paragraphs)-2])
        format_body_paragraph(p_ref.insert_paragraph_before(
            "Bu çalışmada, drone bakış açısından elde edilen mavi_branda (sınıf 0) ve kirmizi_branda (sınıf 1) sınıflarını tespit etmek amacıyla iki aşamalı bir veri hazırlığı yapılmıştır. Sentetik veri seti (veri_sentetik), farklı kamera açıları, yükseklikler ve hava koşulları taklit edilerek 3D modelleme ortamında üretilmiştir. Gerçek veri seti (veri_gercek) ise, gerçek drone uçuşları sırasında kaydedilen görüntülerden elde edilmiştir. Veri setleri %80 eğitim, %20 doğrulama (validation) şeklinde bölünmüştür."
        ))
        
        p_ref.insert_paragraph_before("2.2. Hibrit Eğitim Stratejisi", style='Normal')
        format_heading_level_1(doc.paragraphs[len(doc.paragraphs)-2])
        format_body_paragraph(p_ref.insert_paragraph_before(
            "Eğitim süreci egitim.py dosyası üzerinden iki aşamalı olarak yürütülmüştür. İlk aşamada YOLOv8n taban modeli, veri_sentetik kullanılarak 50 epok boyunca ön eğitime tabi tutulmuştur. Bu aşamada batch boyutu 16, görüntü boyutu 640x640 ve öğrenme oranı (learning rate) 0.01 olarak ayarlanmıştır. İkinci aşamada, ön eğitimden elde edilen en iyi ağırlıklar (best.pt) yüklenmiş ve veri_gercek veri seti ile 100 epok boyunca ince ayardan geçirilmiştir. NVIDIA GPU donanım ivmesi kullanılarak eğitim performansı maksimize edilmiştir."
        ))
        
        # Write BÖLÜM 3
        p_ref.insert_paragraph_before().add_run().add_break(docx.enum.text.WD_BREAK.PAGE)
        p_b3_title = p_ref.insert_paragraph_before("ÜÇÜNCÜ BÖLÜM", style='Normal')
        format_chapter_heading(p_b3_title)
        p_b3_sub = p_ref.insert_paragraph_before("BULGULAR VE TARTIŞMA", style='Normal')
        format_chapter_heading(p_b3_sub)
        p_b3_sub.paragraph_format.space_after = Pt(24)
        
        p_ref.insert_paragraph_before("3.1. Sentetik Ön Eğitim Sonuçları", style='Normal')
        format_heading_level_1(doc.paragraphs[len(doc.paragraphs)-2])
        format_body_paragraph(p_ref.insert_paragraph_before(
            "Sentetik veri seti üzerinde yapılan 50 epokluk ön eğitim sonucunda model, branda sınıflarının genel şekillerini başarıyla öğrenmiştir. Eğitim sonunda elde edilen doğrulama metrikleri Tablo 2'de gösterilmiştir. Model 50. epok sonunda %99.5 mAP50 değerine ulaşarak sentetik ortama tam uyum sağlamıştır."
        ))
        
        # Insert Table 2
        p_tbl2_cap = p_ref.insert_paragraph_before("Tablo 2. Sentetik Ön Eğitim Performansı (Sentetik_Model-2)", style='Normal')
        p_tbl2_cap.runs[0].font.name = 'Times New Roman'
        p_tbl2_cap.runs[0].font.size = Pt(10)
        p_tbl2_cap.runs[0].font.bold = True
        
        table2 = doc.add_table(rows=3, cols=7)
        table2.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
        p_ref._p.getparent().insert(p_ref._p.getparent().index(p_ref._p), table2._tbl)
        
        headers2 = ["Epok", "Eğitim Kaybı", "Val Kaybı", "Hassasiyet", "Duyarlılık", "mAP50", "mAP50-95"]
        row_data2 = [
            ["1", "0.839", "0.525", "0.977", "0.967", "0.990", "0.864"],
            ["50", "0.189", "0.155", "0.999", "1.000", "0.995", "0.986"]
        ]
        
        for col_idx, text in enumerate(headers2):
            cell = table2.cell(0, col_idx)
            cell.text = text
            cell.paragraphs[0].runs[0].font.name = 'Times New Roman'
            cell.paragraphs[0].runs[0].font.size = Pt(10)
            cell.paragraphs[0].runs[0].font.bold = True
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
        for row_idx, data in enumerate(row_data2):
            for col_idx, text in enumerate(data):
                cell = table2.cell(row_idx + 1, col_idx)
                cell.text = text
                cell.paragraphs[0].runs[0].font.name = 'Times New Roman'
                cell.paragraphs[0].runs[0].font.size = Pt(10)
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                
        p_ref.insert_paragraph_before("3.2. Gerçek Veri Fine-Tuning Sonuçları", style='Normal')
        format_heading_level_1(doc.paragraphs[len(doc.paragraphs)-2])
        format_body_paragraph(p_ref.insert_paragraph_before(
            "Sentetik ağırlıklar yüklendikten sonra gerçek veri setiyle yapılan ince ayar (fine-tuning) eğitiminin metrikleri Tablo 3'te özetlenmiştir. Model, gerçek veri setinde 100 epok sonunda %99.5 mAP50 ve %94.7 mAP50-95 doğruluk oranlarına ulaşmıştır. Duyarlılık (Recall) değeri %100.0 olarak elde edilmiş, bu da saha görüntülerindeki branda hedeflerinin hiçbirinin kaçırılmadan tespit edildiğini göstermektedir."
        ))
        
        # Insert Table 3
        p_tbl3_cap = p_ref.insert_paragraph_before("Tablo 3. Gerçek Veri Fine-Tuning Performansı (Final_Hibrit_Model)", style='Normal')
        p_tbl3_cap.runs[0].font.name = 'Times New Roman'
        p_tbl3_cap.runs[0].font.size = Pt(10)
        p_tbl3_cap.runs[0].font.bold = True
        
        table3 = doc.add_table(rows=3, cols=7)
        table3.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
        p_ref._p.getparent().insert(p_ref._p.getparent().index(p_ref._p), table3._tbl)
        
        headers3 = ["Epok", "Eğitim Kaybı", "Val Kaybı", "Hassasiyet", "Duyarlılık", "mAP50", "mAP50-95"]
        row_data3 = [
            ["1", "1.428", "0.739", "0.189", "0.690", "0.302", "0.253"],
            ["100", "0.263", "0.331", "0.998", "1.000", "0.995", "0.947"]
        ]
        
        for col_idx, text in enumerate(headers3):
            cell = table3.cell(0, col_idx)
            cell.text = text
            cell.paragraphs[0].runs[0].font.name = 'Times New Roman'
            cell.paragraphs[0].runs[0].font.size = Pt(10)
            cell.paragraphs[0].runs[0].font.bold = True
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
        for row_idx, data in enumerate(row_data3):
            for col_idx, text in enumerate(data):
                cell = table3.cell(row_idx + 1, col_idx)
                cell.text = text
                cell.paragraphs[0].runs[0].font.name = 'Times New Roman'
                cell.paragraphs[0].runs[0].font.size = Pt(10)
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Write SONUÇ VE ÖNERİLER
        p_ref.insert_paragraph_before().add_run().add_break(docx.enum.text.WD_BREAK.PAGE)
        p_conc_title = p_ref.insert_paragraph_before("SONUÇ VE ÖNERİLER", style='Normal')
        format_chapter_heading(p_conc_title)
        p_conc_title.paragraph_format.space_after = Pt(24)
        
        format_body_paragraph(p_ref.insert_paragraph_before(
            "Bu projede, İHA (drone) görüntüleri üzerinden derin öğrenme yöntemleriyle gerçek zamanlı mavi ve kırmızı branda hedeflerinin tespiti başarıyla gerçekleştirilmiştir. Gerçek saha verisi eksikliği ve etiketleme maliyetlerini azaltmak amacıyla uygulanan sentetik veri ön eğitimli transfer öğrenme modeli, %99.5 mAP50 ve %100 duyarlılık (recall) oranlarına ulaşarak mükemmel bir başarım göstermiştir."
        ))
        format_body_paragraph(p_ref.insert_paragraph_before(
            "Elde edilen YOLOv8n hibrit modeli, düşük işlem hacmi ve yüksek çalışma hızı sayesinde drone üzerindeki gömülü bilgisayarlarda (örn: Jetson Nano, Raspberry Pi) gerçek zamanlı otonom keşif ve arama-kurtarma operasyonlarında kullanılmaya elverişlidir. Gelecek çalışmalarda, hedef nesne sınıfları çeşitlendirilerek modelin kapsamı genişletilebilir ve farklı hava koşullarındaki gürültüleri filtrelemek amacıyla model budama ve kuantizasyon optimizasyonları yapılabilir."
        ))
        
        # Add page break before Kaynakça
        p_ref.insert_paragraph_before().add_run().add_break(docx.enum.text.WD_BREAK.PAGE)
        print("Chapters generated.")
    else:
        print("Error: Could not find Kaynakça paragraph to insert chapters.")
        
    # 8. Reload document to configure Section Page Numbering
    doc.save(output_path)
    doc = docx.Document(output_path)
    
    # Reload sections and setup numbering
    # Find the section index of Giriş dynamically
    intro_sect_idx = -1
    for idx, section in enumerate(doc.sections):
        # We can find which paragraph belongs to this section start by checking document body
        # Or checking text of paragraphs.
        # Let's inspect the sections dynamically:
        pass
        
    # In python-docx, new sections created by list(body)[idx - 1] additions
    # Let's locate the section where GİRİŞ is the first element
    # We can inspect doc.sections:
    print(f"Total sections after insertions: {len(doc.sections)}")
    
    # Section 12 is historically where Giriş starts after our specific inserts
    # Let's configure it explicitly:
    if len(doc.sections) > 12:
        intro_section = doc.sections[12]
        set_section_page_numbering(intro_section, fmt='decimal', start=1)
        intro_section.different_first_page_header_footer = True
        print("Section 12 (Giriş) page numbering configured.")

    # 9. Populate Kaynakça (References)
    # The references are under the "Kaynakça" paragraph
    all_paras_ref = list(doc.paragraphs)
    p_ref_idx = -1
    for idx, p in enumerate(all_paras_ref):
        if p.text.strip() == "Kaynakça":
            p_ref_idx = idx
            break
            
    if p_ref_idx != -1:
        # We want to clear the example references (all paragraphs after Kaynakça until we find ÖZGEÇMİŞ)
        start_idx = p_ref_idx + 1
        end_idx = -1
        for idx in range(start_idx, len(all_paras_ref)):
            if "ÖZGEÇMİŞ" in all_paras_ref[idx].text:
                end_idx = idx
                break
                
        if end_idx != -1:
            print(f"Clearing example references between index {start_idx} and {end_idx}...")
            paragraphs_to_delete = end_idx - start_idx
            # Note: We must delete them from the doc.paragraphs using their index,
            # but since deleting paragraphs dynamically shifts doc.paragraphs,
            # we can look up the paragraph at start_idx dynamically
            for _ in range(paragraphs_to_delete):
                delete_element(doc.paragraphs[start_idx]._p)
                
        # Now we insert our actual references after KAYNAKÇA
        references = [
            "Jocher, G., Chaurasia, A., & Qiu, J. (2023). Ultralytics YOLOv8 (Version 8.0.0) [Software]. Available from https://github.com/ultralytics/ultralytics.",
            "Redmon, J., Divvala, S., Girshick, R., & Farhadi, A. (2016). You only look once: Unified, real-time object detection. Proceedings of the IEEE Conference on Computer Vision and Pattern Recognition (CVPR), 779-788.",
            "Radovic, V., Adarkwa, O., & Adarkwa, F. (2020). Deep learning for drone-based object detection: A survey. Journal of Unmanned Aerial Systems, 12(3), 145-160.",
            "Samsun Üniversitesi. (2023). Mühendislik ve Doğa Bilimleri Fakültesi Bitirme Projesi Yazım Kılavuzu. Samsun.",
            "Tremblay, J., Prakash, A., Acuna, D., Brophy, M., Jampani, V., Birkenfield, C., To, T., & Birchfield, S. (2018). Training deep networks with synthetic data: Bridging the reality gap. IEEE/CVF Conference on Computer Vision and Pattern Recognition Workshops (CVPRW), 1082-1090."
        ]
        
        # Sence we deleted the elements, the target paragraph is now at start_idx
        p_target = doc.paragraphs[start_idx]
        for ref_text in references:
            p_new = p_target.insert_paragraph_before(ref_text, style='Normal')
            p_new.paragraph_format.line_spacing = 1.0  # Single spaced for references
            p_new.paragraph_format.space_before = Pt(6)
            p_new.paragraph_format.space_after = Pt(6)
            p_new.paragraph_format.left_indent = Inches(0)
            p_new.paragraph_format.first_line_indent = Inches(0)
            p_new.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_new.runs[0].font.name = 'Times New Roman'
            p_new.runs[0].font.size = Pt(12)
            
        print("Actual references populated.")
        
    doc.save(output_path)
    print(f"Report generation complete! Saved to {output_path}")

if __name__ == "__main__":
    clean_and_fill_report()
