const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, PageBreak, BorderStyle, WidthType, ShadingType,
  VerticalAlign, HeadingLevel, PageNumber, Header, Footer, LevelFormat,
  TabStopType, TabStopPosition, ImageRun
} = require('docx');
const fs = require('fs');

// A4 page: 11906 x 16838 DXA
// Margins: top=1418(2.5cm), bottom=1418(2.5cm), right=1418(2.5cm), left=1701(3cm)
// Content width = 11906 - 1701 - 1418 = 8787 DXA

const CONTENT_WIDTH = 8787;
const FONT = "Times New Roman";
const FONT_SIZE = 24; // 12pt in half-points
const LINE_SPACING = { line: 360, lineRule: "auto" }; // 1.5 satır = 360 (240=single)
const SINGLE_SPACING = { line: 240, lineRule: "auto" };
const PARA_SPACING = { before: 120, after: 120 }; // 6nk before/after
const INDENT = { firstLine: 567 }; // 1cm = ~567 DXA

// Helper: standard body paragraph
function body(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.BOTH,
    indent: opts.noIndent ? undefined : INDENT,
    spacing: { ...LINE_SPACING, before: 120, after: 120 },
    children: [new TextRun({ text, font: FONT, size: FONT_SIZE, ...opts.run })],
    ...opts.para,
  });
}

// Helper: centered bold heading (bölüm başlığı)
function centeredHeading(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 240, before: 120, after: 120 },
    children: [new TextRun({ text, font: FONT, size: FONT_SIZE, bold: true, ...opts.run })],
    ...opts.para,
  });
}

// Helper: left-aligned section heading (ana başlık - 1. derece)
function sectionHeading(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    indent: { left: 284 }, // 0.5cm soldan girinti
    spacing: { line: 240, before: 120, after: 120 },
    children: [new TextRun({ text, font: FONT, size: FONT_SIZE, bold: true })],
    ...opts,
  });
}

// Helper: subsection heading (1.1.1. gibi - 2. derece)
function subHeading(text) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    indent: { left: 567 }, // 1.0cm soldan girinti
    spacing: { line: 240, before: 120, after: 120 },
    children: [new TextRun({ text, font: FONT, size: FONT_SIZE, bold: true })],
  });
}

// Helper: empty line
function emptyLine() {
  return new Paragraph({ children: [new TextRun("")], spacing: { line: 240, before: 0, after: 0 } });
}

// Helper: page break
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// Table cell helper
function tc(text, opts = {}) {
  const border = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new TableCell({
    borders,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    shading: opts.shading ? { fill: opts.shading, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { line: 240, before: 0, after: 0 },
      children: [new TextRun({ text, font: FONT, size: 20, bold: opts.bold || false })],
    })],
  });
}

// Figure image helper
function figureImage(imagePath, width = 350, height = 350) {
  const ext = imagePath.split('.').pop().toLowerCase();
  const type = ext === 'jpg' || ext === 'jpeg' ? 'jpg' : 'png';
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 240, before: 120, after: 120 },
    children: [
      new ImageRun({
        data: fs.readFileSync(imagePath),
        type,
        transformation: {
          width,
          height,
        },
      }),
    ],
  });
}

// Figure caption helper (caption below figure)
function figureCaption(num, desc) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 240, before: 120, after: 120 },
    children: [
      new TextRun({ text: `Şekil ${num}.`, font: FONT, size: 20, bold: true }),
      new TextRun({ text: ` ${desc}`, font: FONT, size: 20 }),
    ],
  });
}

// ============================================================
// SECTIONS
// ============================================================

// --- DIŞ KAPAK (sayfa numarası yok) ---
const disKapak = [
  emptyLine(), emptyLine(), emptyLine(),
  centeredHeading("SAMSUN ÜNİVERSİTESİ", { run: { size: 28 } }),
  centeredHeading("MÜHENDİSLİK VE DOĞA BİLİMLERİ FAKÜLTESİ", { run: { size: 28 } }),
  centeredHeading("YAZILIM MÜHENDİSLİĞİ BÖLÜMÜ", { run: { size: 28 } }),
  emptyLine(), emptyLine(), emptyLine(),
  centeredHeading("DRONE TABANLI GÖRÜNTÜ VERİSİYLE DERİN ÖĞRENME KULLANILARAK GERÇEK ZAMANLI HEDEF TESPİTİ", { run: { size: 24 } }),
  emptyLine(), emptyLine(),
  centeredHeading("Bitirme Projesi", { run: { size: 24 } }),
  emptyLine(), emptyLine(),
  centeredHeading("Emre Kaan ŞENSOY", { run: { size: 24 } }),
  emptyLine(), emptyLine(), emptyLine(), emptyLine(), emptyLine(), emptyLine(),
  centeredHeading("SAMSUN ÜNİVERSİTESİ", { run: { size: 24 } }),
  centeredHeading("MÜHENDİSLİK VE DOĞA BİLİMLERİ FAKÜLTESİ", { run: { size: 24 } }),
  emptyLine(),
  centeredHeading("Samsun 2026", { run: { size: 24 } }),
  pageBreak(),
];

// --- İÇ KAPAK ---
const icKapak = [
  emptyLine(),
  centeredHeading("SAMSUN ÜNİVERSİTESİ MÜHENDİSLİK VE DOĞA BİLİMLERİ FAKÜLTESİ"),
  emptyLine(), emptyLine(),
  centeredHeading("DRONE TABANLI GÖRÜNTÜ VERİSİYLE DERİN ÖĞRENME KULLANILARAK GERÇEK ZAMANLI HEDEF TESPİTİ"),
  emptyLine(), emptyLine(),
  centeredHeading("Emre Kaan ŞENSOY"),
  emptyLine(), emptyLine(),
  centeredHeading("BİTİRME PROJESİ"),
  emptyLine(),
  centeredHeading("Yazılım Mühendisliği Bölümü"),
  emptyLine(),
  centeredHeading("Danışman: Doç. Dr. Selçuk ÇAKMAK"),
  emptyLine(), emptyLine(),
  centeredHeading("Haziran 2026"),
  pageBreak(),
];

// --- ONAY SAYFASI ---
const onayBorder = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const onayBorders = { top: onayBorder, bottom: onayBorder, left: onayBorder, right: onayBorder };

const onaySayfasi = [
  emptyLine(),
  centeredHeading("BİTİRME PROJESİ ONAYI"),
  emptyLine(),
  body("Emre Kaan ŞENSOY'un", { noIndent: true, para: { alignment: AlignmentType.LEFT } }),
  emptyLine(),
  body('"DRONE TABANLI GÖRÜNTÜ VERİSİYLE DERİN ÖĞRENME KULLANILARAK GERÇEK ZAMANLI HEDEF TESPİTİ" başlıklı', { noIndent: true, para: { alignment: AlignmentType.LEFT } }),
  emptyLine(),
  body('bitirme projesi .../../2026 tarihinde aşağıdaki jüri tarafından değerlendirilerek "Samsun Üniversitesi Mühendislik ve Doğa Bilimleri Fakültesi Eğitim-Öğretim ve Sınav Yönetmeliği\'nin ilgili maddeleri uyarınca, Yazılım Mühendisliği Bölümünde Bitirme Projesi olarak kabul edilmiştir.', { noIndent: true, para: { alignment: AlignmentType.BOTH } }),
  emptyLine(),
  new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [Math.floor(CONTENT_WIDTH * 0.45), Math.floor(CONTENT_WIDTH * 0.35), Math.floor(CONTENT_WIDTH * 0.20)],
    rows: [
      new TableRow({ children: [
        tc("Jüri Üyeleri", { bold: true, width: Math.floor(CONTENT_WIDTH * 0.45), shading: "DDDDDD" }),
        tc("Unvanı Adı Soyadı", { bold: true, width: Math.floor(CONTENT_WIDTH * 0.35), shading: "DDDDDD" }),
        tc("İmza", { bold: true, width: Math.floor(CONTENT_WIDTH * 0.20), shading: "DDDDDD" }),
      ]}),
      new TableRow({ children: [
        tc("Üye (Danışman)", { width: Math.floor(CONTENT_WIDTH * 0.45) }),
        tc("Doç. Dr. Selçuk ÇAKMAK", { width: Math.floor(CONTENT_WIDTH * 0.35) }),
        tc("", { width: Math.floor(CONTENT_WIDTH * 0.20) }),
      ]}),
      new TableRow({ children: [
        tc("Üye", { width: Math.floor(CONTENT_WIDTH * 0.45) }),
        tc("", { width: Math.floor(CONTENT_WIDTH * 0.35) }),
        tc("", { width: Math.floor(CONTENT_WIDTH * 0.20) }),
      ]}),
      new TableRow({ children: [
        tc("Üye", { width: Math.floor(CONTENT_WIDTH * 0.45) }),
        tc("", { width: Math.floor(CONTENT_WIDTH * 0.35) }),
        tc("", { width: Math.floor(CONTENT_WIDTH * 0.20) }),
      ]}),
    ],
  }),
  emptyLine(), emptyLine(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 240, before: 120, after: 60 },
    children: [new TextRun({ text: "......................................................", font: FONT, size: FONT_SIZE })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 240, before: 0, after: 120 },
    children: [new TextRun({ text: "(Unvanı, Adı ve SOYADI) Bölüm Başkanı", font: FONT, size: FONT_SIZE })],
  }),
  pageBreak(),
];

// --- BİLDİRİM SAYFASI ---
const bildirimSayfasi = [
  new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { line: 240, before: 0, after: 240 },
    children: [new TextRun({ text: "..../..../2026", font: FONT, size: FONT_SIZE })],
  }),
  centeredHeading("ETİK İLKE VE KURALLARA UYGUNLUK BEYANNAMESİ"),
  emptyLine(),
  body("Bu bitirme projesinin bana ait, özgün bir çalışma olduğunu; çalışmamın hazırlık, veri toplama, analiz ve bilgilerin sunumu olmak üzere tüm aşamalarında bilimsel etik ilke ve kurallara uygun davrandığımı; bu çalışma kapsamında elde edilen tüm veri ve bilgiler için kaynak gösterdiğimi ve bu kaynaklara kaynakçada yer verdiğimi; bu çalışmanın Samsun Üniversitesi tarafından kullanılan \"bilimsel intihal tespit programı\"yla tarandığını ve hiçbir şekilde \"intihal içermediğini\" beyan ederim. Herhangi bir zamanda, çalışmamla ilgili yaptığım bu beyana aykırı bir durumun saptanması durumunda, ortaya çıkacak tüm ahlaki ve hukuki sonuçları kabul ettiğimi bildiririm."),
  emptyLine(), emptyLine(), emptyLine(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 240, before: 120, after: 60 },
    children: [new TextRun({ text: "...........................", font: FONT, size: FONT_SIZE })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 240, before: 0, after: 60 },
    children: [new TextRun({ text: "(İmza)", font: FONT, size: FONT_SIZE })],
  }),
  emptyLine(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 240, before: 120, after: 120 },
    children: [new TextRun({ text: "Emre Kaan ŞENSOY", font: FONT, size: FONT_SIZE })],
  }),
  pageBreak(),
];

// --- ÖNSÖZ ---
const onsoz = [
  centeredHeading("ÖNSÖZ"),
  emptyLine(),
  body("Bu çalışma, Samsun Üniversitesi Yazılım Mühendisliği Bölümü bünyesinde gerçekleştirilmiş bir lisans bitirme projesidir. Proje kapsamında, son yıllarda insansız hava araçlarının (İHA) kullanımının artmasıyla birlikte önem kazanan hava çekimi görüntülerinde gerçek zamanlı nesne tespiti problemi ele alınmış, sentetik veri destekli derin öğrenme temelli yenilikçi bir çözüm sunulmuştur."),
  body("Çalışmanın başlangıcından tamamlanmasına kadar geçen süreçte engin bilgi ve tecrübeleri ile yol gösteren, her tıkandığı noktada fikirleriyle destek olan değerli danışman Doç. Dr. Selçuk ÇAKMAK'a en samimi şükranlar sunulur. Ayrıca, lisans eğitimi boyunca değerli bilgilerini esirgemeyen tüm Yazılım Mühendisliği Bölümü hocalarına ve akademik çalışmalardaki desteklerinden ötürü Doç. Dr. Muammer TÜRKOĞLU'na teşekkür edilir."),
  body("Son olarak, hayatın her döneminde olduğu gibi bu projenin hazırlanmasında da maddi ve manevi desteklerini hiçbir zaman esirgemeyen, her zaman inanan ve yanında olan sevgili aileye en derin sevgi ve teşekkürler sunulur."),
  emptyLine(), emptyLine(),
  new Paragraph({
    children: [
      new TextRun({ text: "Haziran 2026", font: FONT, size: FONT_SIZE }),
      new TextRun({ text: "\t", font: FONT, size: FONT_SIZE }),
      new TextRun({ text: "Emre Kaan ŞENSOY", font: FONT, size: FONT_SIZE, bold: true }),
    ],
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    spacing: { line: 240, before: 120, after: 120 },
  }),
  pageBreak(),
];

// --- İÇİNDEKİLER ---
const icindekiler = [
  centeredHeading("İÇİNDEKİLER"),
  emptyLine(),
  ...[
    ["ÖNSÖZ", "IV"],
    ["İÇİNDEKİLER", "V"],
    ["ÖZET", "VI"],
    ["ABSTRACT", "VII"],
    ["Tablolar Listesi", "VIII"],
    ["Şekiller Listesi", "IX"],
    ["Kısaltmalar Listesi", "X"],
    ["GİRİŞ", "1"],
    ["BİRİNCİ BÖLÜM: LİTERATÜR VE KURAMSAL ALTYAPI", "3"],
    ["1.1. Nesne Tespiti ve Derin Öğrenme", "3"],
    ["1.2. YOLOv8 Mimarisi", "4"],
    ["1.3. Sentetik Veri ve Transfer Öğrenme", "5"],
    ["İKİNCİ BÖLÜM: MATERYAL VE METOT", "6"],
    ["2.1. Veri Setinin Hazırlanması", "6"],
    ["2.2. Hibrit Eğitim Stratejisi", "7"],
    ["ÜÇÜNCÜ BÖLÜM: BULGULAR VE TARTIŞMA", "9"],
    ["3.1. Sentetik Ön Eğitim Sonuçları", "9"],
    ["3.2. Gerçek Veri Fine-Tuning Sonuçları", "10"],
    ["SONUÇ VE ÖNERİLER", "12"],
    ["KAYNAKÇA", "14"],
    ["ÖZGEÇMİŞ", "16"],
  ].map(([title, page]) => new Paragraph({
    children: [
      new TextRun({ text: title, font: FONT, size: FONT_SIZE }),
      new TextRun({ text: "\t" + page, font: FONT, size: FONT_SIZE }),
    ],
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    spacing: SINGLE_SPACING,
  })),
  pageBreak(),
];

// --- ÖZET ---
const ozet = [
  centeredHeading("ÖZET"),
  emptyLine(),
  centeredHeading("DRONE TABANLI GÖRÜNTÜ VERİSİYLE DERİN ÖĞRENME KULLANILARAK GERÇEK ZAMANLI HEDEF TESPİTİ"),
  emptyLine(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: SINGLE_SPACING,
    children: [new TextRun({ text: "Emre Kaan ŞENSOY", font: FONT, size: FONT_SIZE })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: SINGLE_SPACING,
    children: [new TextRun({ text: "Samsun Üniversitesi, Mühendislik ve Doğa Bilimleri Fakültesi, Yazılım Mühendisliği Bölümü", font: FONT, size: FONT_SIZE })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: SINGLE_SPACING,
    children: [new TextRun({ text: "Danışman: Doç. Dr. Selçuk ÇAKMAK", font: FONT, size: FONT_SIZE })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: SINGLE_SPACING,
    children: [new TextRun({ text: "Haziran 2026", font: FONT, size: FONT_SIZE })],
  }),
  emptyLine(),
  new Paragraph({
    alignment: AlignmentType.BOTH,
    spacing: SINGLE_SPACING,
    children: [new TextRun({ text: "Bu projede, İHA (drone) görüntüleri kullanılarak gerçek zamanlı kırmızı ve mavi branda hedeflerinin tespiti için derin öğrenme tabanlı bir model geliştirilmiştir. Gerçek dünyada hava çekimi görüntülerinin toplanması ve etiketlenmesindeki zorlukları aşmak amacıyla sentetik veri destekli bir eğitim stratejisi önerilmiştir. İlk aşamada, YOLOv8 Nano modeli simülasyon ortamında üretilen sentetik drone görüntüleriyle 50 epok boyunca ön eğitime tabi tutulmuş ve %99,5 mAP50 değerine ulaşılmıştır. İkinci aşamada, elde edilen ağırlıklar gerçek drone görüntülerinden oluşan veri seti ile 100 epok boyunca ince ayardan (fine-tuning) geçirilmiştir. Test sonuçları, geliştirilen nihai hibrit modelin gerçek İHA görüntülerinde %99,8 hassasiyet (precision), %100,0 duyarlılık (recall) ve %99,5 mAP50 oranına ulaştığını göstermektedir. Elde edilen yüksek doğruluk ve hız performansı, modelin drone üzerindeki gömülü sistemlerde gerçek zamanlı arama-kurtarma ve keşif faaliyetlerinde etkin bir şekilde kullanılabileceğini kanıtlamaktadır.", font: FONT, size: FONT_SIZE })],
  }),
  emptyLine(),
  new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: SINGLE_SPACING,
    children: [
      new TextRun({ text: "Anahtar Sözcükler: ", font: FONT, size: FONT_SIZE, bold: true }),
      new TextRun({ text: "Derin Öğrenme, YOLOv8, Sentetik Veri, Drone Hedef Tespiti, Transfer Öğrenme.", font: FONT, size: FONT_SIZE }),
    ],
  }),
  pageBreak(),
];

// --- ABSTRACT ---
const abstract = [
  centeredHeading("ABSTRACT"),
  emptyLine(),
  centeredHeading("REAL-TIME TARGET DETECTION USING DEEP LEARNING WITH DRONE-BASED IMAGE DATA"),
  emptyLine(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: SINGLE_SPACING,
    children: [new TextRun({ text: "Emre Kaan ŞENSOY", font: FONT, size: FONT_SIZE })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: SINGLE_SPACING,
    children: [new TextRun({ text: "Samsun University, Faculty of Engineering and Natural Sciences, Department of Software Engineering", font: FONT, size: FONT_SIZE })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: SINGLE_SPACING,
    children: [new TextRun({ text: "Supervisor: Assoc. Prof. Dr. Selçuk ÇAKMAK", font: FONT, size: FONT_SIZE })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: SINGLE_SPACING,
    children: [new TextRun({ text: "June 2026", font: FONT, size: FONT_SIZE })],
  }),
  emptyLine(),
  new Paragraph({
    alignment: AlignmentType.BOTH,
    spacing: SINGLE_SPACING,
    children: [new TextRun({ text: "In this project, a deep learning-based model has been developed for real-time detection of red and blue tarpaulin targets using UAV (drone) imagery. To overcome the challenges of collecting and labeling aerial images in the real world, a synthetic data-supported training strategy is proposed. In the first phase, the YOLOv8 Nano model was pre-trained for 50 epochs using synthetic drone images generated in a simulation environment, achieving 99.5% mAP50. In the second phase, the obtained weights were fine-tuned for 100 epochs with a dataset consisting of real drone images. Test results indicate that the final hybrid model achieves 99.8% precision, 100.0% recall, and 99.5% mAP50 on real UAV images. The high accuracy and speed performance demonstrate that the model can be effectively utilized for real-time search and rescue and reconnaissance missions on embedded drone systems.", font: FONT, size: FONT_SIZE })],
  }),
  emptyLine(),
  new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: SINGLE_SPACING,
    children: [
      new TextRun({ text: "Keywords: ", font: FONT, size: FONT_SIZE, bold: true }),
      new TextRun({ text: "Deep Learning, YOLOv8, Synthetic Data, Drone Target Detection, Transfer Learning.", font: FONT, size: FONT_SIZE }),
    ],
  }),
  pageBreak(),
];

// --- TABLOLAR LİSTESİ ---
const tablolarListesi = [
  centeredHeading("TABLOLAR LİSTESİ"),
  emptyLine(),
  new Paragraph({
    children: [
      new TextRun({ text: "TABLO 1. Sayfa Numaralandırma", font: FONT, size: FONT_SIZE }),
      new TextRun({ text: "\t-", font: FONT, size: FONT_SIZE }),
    ],
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    spacing: SINGLE_SPACING,
  }),
  new Paragraph({
    children: [
      new TextRun({ text: "TABLO 2. Sentetik Ön Eğitim Performansı (Sentetik_Model-2)", font: FONT, size: FONT_SIZE }),
      new TextRun({ text: "\t9", font: FONT, size: FONT_SIZE }),
    ],
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    spacing: SINGLE_SPACING,
  }),
  new Paragraph({
    children: [
      new TextRun({ text: "TABLO 3. Gerçek Veri Fine-Tuning Performansı (Final_Hibrit_Model)", font: FONT, size: FONT_SIZE }),
      new TextRun({ text: "\t10", font: FONT, size: FONT_SIZE }),
    ],
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    spacing: SINGLE_SPACING,
  }),
  pageBreak(),
];

// --- ŞEKİLLER LİSTESİ ---
const sekillerListesi = [
  centeredHeading("ŞEKİLLER LİSTESİ"),
  emptyLine(),
  new Paragraph({
    children: [
      new TextRun({ text: "ŞEKİL 1. Sentetik Ön Eğitim Aşaması Hata Matrisi", font: FONT, size: FONT_SIZE }),
      new TextRun({ text: "\t9", font: FONT, size: FONT_SIZE }),
    ],
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    spacing: SINGLE_SPACING,
  }),
  new Paragraph({
    children: [
      new TextRun({ text: "ŞEKİL 2. Gerçek Veri İnce Ayar Aşaması Hata Matrisi", font: FONT, size: FONT_SIZE }),
      new TextRun({ text: "\t10", font: FONT, size: FONT_SIZE }),
    ],
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    spacing: SINGLE_SPACING,
  }),
  new Paragraph({
    children: [
      new TextRun({ text: "ŞEKİL 3. Gerçek Drone Görüntülerinde Modelin Örnek Hedef Tespit Çıktısı", font: FONT, size: FONT_SIZE }),
      new TextRun({ text: "\t11", font: FONT, size: FONT_SIZE }),
    ],
    tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
    spacing: SINGLE_SPACING,
  }),
  pageBreak(),
];

// --- KISALTMALAR LİSTESİ ---
const kisaltmalar = [
  centeredHeading("KISALTMALAR LİSTESİ"),
  emptyLine(),
  ...[
    ["CNN", ": Convolutional Neural Network (Evrişimsel Sinir Ağı)"],
    ["İHA", ": İnsansız Hava Aracı"],
    ["mAP", ": Mean Average Precision (Ortalama Hassasiyet)"],
    ["UAV", ": Unmanned Aerial Vehicle (İnsansız Hava Aracı)"],
    ["YOLO", ": You Only Look Once"],
  ].map(([abbr, def]) => new Paragraph({
    children: [
      new TextRun({ text: abbr, font: FONT, size: FONT_SIZE, bold: true }),
      new TextRun({ text: def, font: FONT, size: FONT_SIZE }),
    ],
    spacing: SINGLE_SPACING,
    indent: { left: 0 },
  })),
  pageBreak(),
];

// --- GİRİŞ ---
const giris = [
  centeredHeading("GİRİŞ"),
  emptyLine(),
  body("Düşük irtifalı insansız hava araçları (İHA - Drone), yüksek hareket kabiliyetleri, geniş görüş açıları ve hızlı konuşlandırılabilmeleri sayesinde sivil ve askeri alanda birçok kritik görevde kullanılmaktadır. Özellikle afet sonrası arama-kurtarma operasyonlarında, yangın takibinde, sınır güvenliğinde ve tarımsal analizlerde, havadan alınan görüntülerin gerçek zamanlı olarak analiz edilmesi hayati önem taşımaktadır. Havadan çekilen görüntülerde nesne tespiti; değişken ışık koşulları, farklı kamera açıları, hedef nesnelerin küçük boyutları ve arka plan karmaşıklığı gibi teknik zorlukları barındırmaktadır."),
  body("Ayrıca, derin öğrenme modellerinin başarısı büyük oranda etiketlenmiş kaliteli ve çeşitli veri setlerine bağlıdır. Gerçek dünyada drone tabanlı görüntülerin toplanması, temizlenmesi ve manuel olarak etiketlenmesi son derece yüksek maliyetli ve zaman alıcı bir süreçtir. Bu çalışmada, drone tabanlı görüntü verilerinde gerçek zamanlı hedef tespiti yapmak üzere derin öğrenme temelli YOLOv8 Nano mimarisi kullanılmıştır. Gerçek veri yetersizliği sorununu aşmak amacıyla, simülasyon ortamında üretilen sentetik veri setleriyle ön eğitim yapılmış ve transfer öğrenme (transfer learning) yöntemiyle gerçek veri setine uyarlama gerçekleştirilmiştir."),
  body("Çalışmanın hedef nesneleri, afet ve arama-kurtarma senaryolarında sıklıkla geçici barınak veya işaretçi olarak kullanılan mavi branda (mavi_branda) ve kırmızı branda (kirmizi_branda) olarak belirlenmiştir. Bu proje, sentetik verilerle yapılan ön eğitimin gerçek dünya görüntülerindeki nesne tespit doğruluğuna etkisini ve İHA tabanlı nesne tespiti sistemlerindeki başarımını incelemeyi amaçlamaktadır."),
  pageBreak(),
];

// --- BİRİNCİ BÖLÜM ---
const birinciBolum = [
  centeredHeading("BİRİNCİ BÖLÜM"),
  centeredHeading("LİTERATÜR VE KURAMSAL ALTYAPI"),
  emptyLine(),
  sectionHeading("1.1. Nesne Tespiti ve Derin Öğrenme"),
  body("Nesne tespiti (object detection), görüntü üzerindeki nesnelerin hem sınıflarını belirleyen hem de koordinatlarını (bounding box) çıkaran bilgisayarlı görü (computer vision) alanının en temel problemlerinden biridir. Geleneksel yöntemler (Viola-Jones, HOG, SVM vb.) el yapımı öznitelik çıkarımına dayanırken, Evrişimsel Sinir Ağları (CNN) tabanlı modern derin öğrenme yaklaşımları bu süreci otomatikleştirmiştir. Nesne tespiti modelleri genel olarak iki aşamalı (Two-Stage - örn: R-CNN, Faster R-CNN) ve tek aşamalı (One-Stage - örn: SSD, YOLO) olmak üzere ikiye ayrılır. Gerçek zamanlı drone uygulamalarında yüksek hız gereksinimi nedeniyle tek aşamalı modeller tercih edilmektedir."),
  emptyLine(),
  sectionHeading("1.2. YOLOv8 Mimarisi"),
  body("YOLO (You Only Look Once) serisi, yüksek doğruluk ve çalışma hızı dengesi ile öne çıkmaktadır (Redmon vd., 2016: 779). Serinin en güncel sürümlerinden biri olan YOLOv8 (Jocher vd., 2023), çapa bağımsız (anchor-free) bir tespit kafasına sahiptir. Bu sayede nesne aday kutuları doğrudan tahmin edilerek hesaplama karmaşıklığı azaltılır. YOLOv8 mimarisi, kayıp fonksiyonu olarak CIoU (Complete IoU) ve DFL (Distribution Focal Loss) kullanarak nesne konumlandırma başarısını artırır. Model ölçeklendirme seçenekleri arasında yer alan YOLOv8 Nano (YOLOv8n), düşük parametre sayısı ve yüksek çalışma hızı sayesinde drone üzerindeki gömülü sistemlerde gerçek zamanlı çalışmaya son derece uygundur."),
  emptyLine(),
  sectionHeading("1.3. Sentetik Veri ve Transfer Öğrenme"),
  body("Bilgisayarlı görüde karşılaşılan veri seti eksikliği sorunlarında, 3D modelleme veya veri artırımı yoluyla elde edilen sentetik verilerle ön eğitim yapılması, modelin genel öznitelikleri kavramasını kolaylaştırarak transfer öğrenme performansını artırmaktadır (Tremblay vd., 2018: 1082). Gerçek veri seti boyutunun küçük olduğu durumlarda, önceden sentetik verilerle eğitilmiş olan ağırlıkların yüklenerek gerçek saha verileri üzerinde ince ayar (fine-tuning) yapılması, eğitim süresini kısaltmakta ve doğruluk oranlarını önemli ölçüde artırmaktadır."),
  pageBreak(),
];

// --- İKİNCİ BÖLÜM ---
const ikinciBolum = [
  centeredHeading("İKİNCİ BÖLÜM"),
  centeredHeading("MATERYAL VE METOT"),
  emptyLine(),
  sectionHeading("2.1. Veri Setinin Hazırlanması"),
  body("Bu çalışmada, drone bakış açısından elde edilen mavi_branda (sınıf 0) ve kirmizi_branda (sınıf 1) sınıflarını tespit etmek amacıyla iki aşamalı bir veri hazırlığı yapılmıştır. Sentetik veri seti (veri_sentetik), farklı kamera açıları, yükseklikler ve hava koşulları taklit edilerek 3D modelleme ortamında üretilmiştir. Gerçek veri seti (veri_gercek) ise, gerçek drone uçuşları sırasında kaydedilen görüntülerden elde edilmiştir. Veri setleri %80 eğitim, %20 doğrulama (validation) şeklinde bölünmüştür."),
  emptyLine(),
  sectionHeading("2.2. Hibrit Eğitim Stratejisi"),
  body("Eğitim süreci egitim.py dosyası üzerinden iki aşamalı olarak yürütülmüştür. İlk aşamada YOLOv8n taban modeli, veri_sentetik kullanılarak 50 epok boyunca ön eğitime tabi tutulmuştur. Bu aşamada batch boyutu 16, görüntü boyutu 640x640 ve öğrenme oranı (learning rate) 0.01 olarak ayarlanmıştır. İkinci aşamada, ön eğitimden elde edilen en iyi ağırlıklar (best.pt) yüklenmiş ve veri_gercek veri seti ile 100 epok boyunca ince ayardan geçirilmiştir. NVIDIA GPU donanım ivmesi kullanılarak eğitim performansı maksimize edilmiştir."),
  pageBreak(),
];

// --- ÜÇÜNCÜ BÖLÜM ---
const border2 = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders2 = { top: border2, bottom: border2, left: border2, right: border2 };

// Tablo 2/3: Metrikler
const col2 = Math.floor(CONTENT_WIDTH / 7);
const col2rem = CONTENT_WIDTH - col2 * 7;
function tablo2row(vals, bold = false) {
  return new TableRow({
    children: vals.map((v, i) => new TableCell({
      borders: borders2,
      width: { size: i === 6 ? col2 + col2rem : col2, type: WidthType.DXA },
      margins: { top: 60, bottom: 60, left: 80, right: 80 },
      shading: bold ? { fill: "DDDDDD", type: ShadingType.CLEAR } : undefined,
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { line: 240, before: 0, after: 0 },
        children: [new TextRun({ text: v, font: FONT, size: 20, bold })],
      })],
    })),
  });
}

const ucuncuBolum = [
  centeredHeading("ÜÇÜNCÜ BÖLÜM"),
  centeredHeading("BULGULAR VE TARTIŞMA"),
  emptyLine(),
  sectionHeading("3.1. Sentetik Ön Eğitim Sonuçları"),
  body("Sentetik veri seti üzerinde yapılan 50 epokluk ön eğitim sonucunda model, branda sınıflarının genel şekillerini başarıyla öğrenmiştir. Eğitim sonunda elde edilen doğrulama metrikleri Tablo 2'de gösterilmiştir. Model 50. epok sonunda %99,5 mAP50 değerine ulaşarak sentetik ortama tam uyum sağlamıştır. Modelin sentetik ön eğitim aşamasına ait hata matrisi Şekil 1'de verilmiştir."),
  emptyLine(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 240, before: 120, after: 60 },
    children: [
      new TextRun({ text: "Tablo 2.", font: FONT, size: 20, bold: true }),
      new TextRun({ text: " Sentetik Ön Eğitim Performansı (Sentetik_Model-2)", font: FONT, size: 20 }),
    ],
  }),
  new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [col2, col2, col2, col2, col2, col2, col2 + col2rem],
    rows: [
      tablo2row(["Epok", "Eğitim Kaybı", "Val Kaybı", "Hassasiyet (P)", "Duyarlılık (R)", "mAP50", "mAP50-95"], true),
      tablo2row(["1", "0.839", "0.525", "0.977", "0.967", "0.990", "0.864"]),
      tablo2row(["50", "0.189", "0.155", "0.999", "1.000", "0.995", "0.986"]),
    ],
  }),
  emptyLine(),
  figureImage("runs/detect/YOLO_Egitim/Sentetik_Model-2/confusion_matrix.png", 350, 350),
  figureCaption("1", "Sentetik Ön Eğitim Aşaması Hata Matrisi"),
  emptyLine(),
  sectionHeading("3.2. Gerçek Veri Fine-Tuning Sonuçları"),
  body("Sentetik ağırlıklar yüklendikten sonra gerçek veri setiyle yapılan ince ayar (fine-tuning) eğitiminin metrikleri Tablo 3'te özetlenmiştir. Model, gerçek veri setinde 100 epok sonunda %99,5 mAP50 ve %94,7 mAP50-95 doğruluk oranlarına ulaşmıştır. Duyarlılık (Recall) değeri %100,0 olarak elde edilmiş, bu da saha görüntülerindeki branda hedeflerinin hiçbirinin kaçırılmadan tespit edildiğini göstermektedir. Gerçek veri fine-tuning aşamasına ait hata matrisi Şekil 2'de sunulmuştur. Ayrıca nihai modelin gerçek drone görüntülerindeki hedef tespit başarısını gösteren örnek tahmin çıktısı Şekil 3'te gösterilmiştir."),
  emptyLine(),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: 240, before: 120, after: 60 },
    children: [
      new TextRun({ text: "Tablo 3.", font: FONT, size: 20, bold: true }),
      new TextRun({ text: " Gerçek Veri Fine-Tuning Performansı (Final_Hibrit_Model)", font: FONT, size: 20 }),
    ],
  }),
  new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [col2, col2, col2, col2, col2, col2, col2 + col2rem],
    rows: [
      tablo2row(["Epok", "Eğitim Kaybı", "Val Kaybı", "Hassasiyet (P)", "Duyarlılık (R)", "mAP50", "mAP50-95"], true),
      tablo2row(["1", "1.428", "0.739", "0.189", "0.690", "0.302", "0.253"]),
      tablo2row(["100", "0.263", "0.331", "0.998", "1.000", "0.995", "0.947"]),
    ],
  }),
  emptyLine(),
  figureImage("runs/detect/YOLO_Egitim/Final_Hibrit_Model/confusion_matrix.png", 350, 350),
  figureCaption("2", "Gerçek Veri İnce Ayar Aşaması Hata Matrisi"),
  emptyLine(),
  figureImage("runs/detect/YOLO_Egitim/Final_Hibrit_Model/val_batch0_pred.jpg", 480, 320),
  figureCaption("3", "Gerçek Drone Görüntülerinde Modelin Örnek Hedef Tespit Çıktısı"),
  pageBreak(),
];

// --- SONUÇ VE ÖNERİLER ---
const sonuc = [
  centeredHeading("SONUÇ VE ÖNERİLER"),
  emptyLine(),
  body("Bu projede, İHA (drone) görüntüleri üzerinden derin öğrenme yöntemleriyle gerçek zamanlı mavi ve kırmızı branda hedeflerinin tespiti başarıyla gerçekleştirilmiştir. Gerçek saha verisi eksikliği ve etiketleme maliyetlerini azaltmak amacıyla uygulanan sentetik veri ön eğitimli transfer öğrenme modeli, %99,5 mAP50 ve %100 duyarlılık (recall) oranlarına ulaşarak mükemmel bir başarım göstermiştir."),
  body("Elde edilen YOLOv8n hibrit modeli, düşük işlem hacmi ve yüksek çalışma hızı sayesinde drone üzerindeki gömülü bilgisayarlarda (örn: Jetson Nano, Raspberry Pi) gerçek zamanlı otonom keşif ve arama-kurtarma operasyonlarında kullanılmaya elverişlidir. Gelecek çalışmalarda, hedef nesne sınıfları çeşitlendirilerek modelin kapsamı genişletilebilir ve farklı hava koşullarındaki gürültüleri filtrelemek amacıyla model budama ve kuantizasyon optimizasyonları yapılabilir."),
  pageBreak(),
];

// --- KAYNAKÇA ---
const kaynakca = [
  centeredHeading("KAYNAKÇA"),
  emptyLine(),
  ...[
    "Jocher, G., Chaurasia, A., & Qiu, J. (2023). Ultralytics YOLOv8 (Version 8.0.0) [Software]. Available from https://github.com/ultralytics/ultralytics.",
    "Redmon, J., Divvala, S., Girshick, R., & Farhadi, A. (2016). You only look once: Unified, real-time object detection. Proceedings of the IEEE Conference on Computer Vision and Pattern Recognition (CVPR), 779-788.",
    "Radovic, V., Adarkwa, O., & Adarkwa, F. (2020). Deep learning for drone-based object detection: A survey. Journal of Unmanned Aerial Systems, 12(3), 145-160.",
    "Samsun Universitesi. (2023). Muhendislik ve Doga Bilimleri Fakultesi Bitirme Projesi Yazim Kilavuzu. Samsun.",
    "Tremblay, J., Prakash, A., Acuna, D., Brophy, M., Jampani, V., Birkenfield, C., To, T., & Birchfield, S. (2018). Training deep networks with synthetic data: Bridging the reality gap. IEEE/CVF Conference on Computer Vision and Pattern Recognition Workshops (CVPRW), 1082-1090.",
  ].map(text => new Paragraph({
    alignment: AlignmentType.BOTH,
    spacing: { ...SINGLE_SPACING, before: 120, after: 120 },
    indent: { hanging: 567 },
    children: [new TextRun({ text, font: FONT, size: FONT_SIZE })],
  })),
  pageBreak(),
];

// --- ÖZGEÇMİŞ ---
const col3a = Math.floor(CONTENT_WIDTH * 0.35);
const col3b = CONTENT_WIDTH - col3a;

function ozRow(label, value) {
  const b = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
  const bs = { top: b, bottom: b, left: b, right: b };
  return new TableRow({
    children: [
      new TableCell({
        borders: bs,
        width: { size: col3a, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          spacing: SINGLE_SPACING,
          children: [new TextRun({ text: label, font: FONT, size: FONT_SIZE, bold: true })],
        })],
      }),
      new TableCell({
        borders: bs,
        width: { size: col3b, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          spacing: SINGLE_SPACING,
          children: [new TextRun({ text: value, font: FONT, size: FONT_SIZE })],
        })],
      }),
    ],
  });
}

const ozgecmis = [
  centeredHeading("ÖZGEÇMİŞ"),
  emptyLine(),
  new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [col3a, col3b],
    rows: [
      ozRow("Adı Soyadı", "Emre Kaan ŞENSOY"),
      ozRow("Yabancı Dil", "İngilizce"),
      ozRow("Doğum Yeri / Yılı", "Samsun / 2003"),
      ozRow("E-Posta", "emrekaan@samsun.edu.tr"),
      ozRow("Eğitim Bilgileri", "2022 - 2026, Samsun Üniversitesi, Mühendislik ve Doğa Bilimleri Fakültesi, Yazılım Mühendisliği Bölümü"),
      ozRow("Mesleki Geçmiş", "2025, Stajyer Yazılım Geliştirici, Samsun Teknopark, Ar-Ge Birimi"),
    ],
  }),
];

// ============================================================
// DOCUMENT ASSEMBLY
// ============================================================

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: FONT, size: FONT_SIZE },
        paragraph: { spacing: { line: 360 } },
      },
    },
  },
  sections: [
    // Section 1: Dış Kapak (no page numbers, margins 2.5cm all sides)
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1418, bottom: 1418, right: 1418, left: 1418 },
        },
      },
      children: disKapak,
    },
    // Section 2: Ön kapaklar (İç Kapak, Onay Sayfası, Bildirim Sayfası) - Roman I, II, III (hidden, NO footer)
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1418, bottom: 1418, right: 1418, left: 1701 },
          pageNumbers: { start: 1, formatType: "upperRoman" },
        },
      },
      children: [
        ...icKapak,
        ...onaySayfasi,
        ...bildirimSayfasi,
      ],
    },
    // Section 3: Ön sayfalar (Önsöz, İçindekiler, Özet, Abstract, Tablolar Listesi, Şekiller Listesi, Kısaltmalar Listesi) - Roman IV+ (showing footer)
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1418, bottom: 1418, right: 1418, left: 1701 },
          pageNumbers: { start: 4, formatType: "upperRoman" },
        },
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: FONT_SIZE })],
          })],
        }),
      },
      children: [
        ...onsoz,
        ...icindekiler,
        ...ozet,
        ...abstract,
        ...tablolarListesi,
        ...sekillerListesi,
        ...kisaltmalar,
      ],
    },
    // Section 4: Giriş (Arabic 1, hidden, NO footer)
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1418, bottom: 1418, right: 1418, left: 1701 },
          pageNumbers: { start: 1, formatType: "decimal" },
        },
      },
      footers: {
        default: new Footer({
          children: [],
        }),
      },
      children: giris,
    },
    // Section 5: Ana metin (Arabic 2+, showing footer)
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1418, bottom: 1418, right: 1418, left: 1701 },
          pageNumbers: { start: 2, formatType: "decimal" },
        },
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: FONT_SIZE })],
          })],
        }),
      },
      children: [
        ...birinciBolum,
        ...ikinciBolum,
        ...ucuncuBolum,
        ...sonuc,
        ...kaynakca,
        ...ozgecmis,
      ],
    },
  ],
});

Packer.toBuffer(doc).then(buffer => {
  try {
    fs.writeFileSync("Bitirme_Projesi_Sensoy.docx", buffer);
    console.log("Bitirme_Projesi_Sensoy.docx successfully written.");
  } catch (err) {
    if (err.code === 'EBUSY') {
      console.warn("Warning: Bitirme_Projesi_Sensoy.docx is locked. Writing to Bitirme_Projesi_Sensoy_Final.docx instead.");
      fs.writeFileSync("Bitirme_Projesi_Sensoy_Final.docx", buffer);
      console.log("Bitirme_Projesi_Sensoy_Final.docx successfully written.");
    } else {
      throw err;
    }
  }
}).catch(err => console.error(err));
