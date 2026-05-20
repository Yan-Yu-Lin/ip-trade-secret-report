const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');

const ROOT = path.resolve(__dirname, '..');
const OUT = path.join(ROOT, 'pptx-native', 'presentation.pptx');
const IMG = (name) => path.join(ROOT, 'ppt', 'images', name);

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Arthur Lin';
pptx.subject = 'Trade Secrets IPR — Technology Licensing';
pptx.title = '營業秘密與技術授權';
pptx.company = 'NTUT';
pptx.lang = 'zh-TW';
pptx.theme = {
  headFontFace: 'Microsoft JhengHei',
  bodyFontFace: 'Microsoft JhengHei',
  lang: 'zh-TW',
};
pptx.defineLayout({ name: 'CUSTOM_WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'CUSTOM_WIDE';
pptx.margin = 0;

const W = 13.333;
const H = 7.5;
const C = {
  paper: 'FAFAF8',
  ink: '0A0A0A',
  accent: '002FA7',
  accentSoft: 'DDE6FF',
  grey: 'F1F1EF',
  grey2: 'E4E4E1',
  text2: '6B6B67',
  white: 'FFFFFF',
};
const F = {
  zh: 'Microsoft JhengHei',
  en: 'Arial',
};

const notesSource = fs.readFileSync(path.join(ROOT, '08-完整連貫敘述-v3.md'), 'utf8');
const noteSections = notesSource.split(/^---$/m).map(s => s.trim()).filter(Boolean);
function noteFor(i, extra = '') {
  const groups = [
    [0], [1, 2], [2], [1, 2], [3], [4], [7], [7], [7], [11],
    [4], [4], [4], [4], [3, 4], [4], [3], [3], [4], [4], [4], [4],
    [8], [11],
  ];
  const body = (groups[i - 1] || [0]).map(idx => noteSections[idx] || '').join('\n\n');
  return `${extra ? extra + '\n\n' : ''}${body}`.replace(/^#.*\n?/gm, '').replace(/\*\*/g, '').trim();
}

function slideBase(slide, meta, num, dark = false) {
  slide.background = { color: dark ? C.ink : C.paper };
  const color = dark ? 'EDEDE8' : C.ink;
  slide.addText(meta.toUpperCase(), {
    x: 0.62, y: 0.42, w: 5.2, h: 0.18, fontFace: F.en, fontSize: 5.5,
    charSpace: 1.6, color, breakLine: false, margin: 0,
  });
  slide.addText(`${num} / 24`, {
    x: 11.8, y: 0.42, w: 0.9, h: 0.18, fontFace: F.en, fontSize: 6,
    charSpace: 1.3, color, align: 'right', margin: 0,
  });
}
function add(slide, text, x, y, w, h, opt = {}) {
  slide.addText(text, {
    x, y, w, h, margin: opt.margin ?? 0.02,
    fontFace: opt.fontFace || F.zh,
    fontSize: opt.fontSize || 16,
    color: opt.color || C.ink,
    bold: !!opt.bold,
    italic: !!opt.italic,
    breakLine: false,
    fit: 'shrink',
    valign: opt.valign || 'top',
    align: opt.align || 'left',
    lineSpacingMultiple: opt.lineSpacingMultiple || 0.9,
    paraSpaceAfterPt: 0,
    charSpace: opt.charSpace || 0,
    rotate: opt.rotate,
  });
}
function meta(slide, text, x, y, w, dark = false, color) {
  add(slide, text.toUpperCase(), x, y, w, 0.22, {
    fontFace: F.en, fontSize: 5.6, color: color || (dark ? 'CFCFCA' : C.text2), charSpace: 1.7,
  });
}
function title(slide, text, x, y, w, h, size = 38, color = C.ink) {
  add(slide, text, x, y, w, h, { fontSize: size, color, lineSpacingMultiple: 0.78 });
}
function body(slide, text, x, y, w, h, size = 13, color = C.text2) {
  add(slide, text, x, y, w, h, { fontSize: size, color, lineSpacingMultiple: 1.05 });
}
function rect(slide, x, y, w, h, fill = C.grey, line = fill, opt = {}) {
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w, h,
    fill: { color: fill, transparency: opt.transparency || 0 },
    line: { color: line, transparency: opt.lineTransparency ?? 100, width: opt.lineWidth || 0.4 },
  });
}
function line(slide, x1, y1, x2, y2, color = C.ink, width = 1, arrow = false, dash) {
  slide.addShape(pptx.ShapeType.line, {
    x: x1, y: y1, w: x2 - x1, h: y2 - y1,
    line: { color, width, beginArrowType: 'none', endArrowType: arrow ? 'triangle' : 'none', dash },
  });
}
function card(slide, x, y, w, h, kicker, head, text, opts = {}) {
  const fill = opts.accent ? C.accent : opts.ink ? C.ink : C.grey;
  const txt = opts.accent || opts.ink ? C.white : C.ink;
  const muted = opts.accent || opts.ink ? 'D7D7D2' : C.text2;
  const compact = h < 1.25;
  rect(slide, x, y, w, h, fill);
  if (kicker) meta(slide, kicker, x + 0.14, y + 0.16, w - 0.28, opts.ink, opts.accent ? C.white : muted);
  add(slide, head, x + 0.14, y + 0.42, w - 0.28, h * 0.34, {
    fontSize: opts.headSize || (compact ? 13.2 : 17), color: txt, lineSpacingMultiple: 0.86,
  });
  if (text) body(slide, text, x + 0.14, compact ? y + 0.74 : y + h - 0.48, w - 0.28, compact ? 0.22 : 0.35, opts.bodySize || (compact ? 7.2 : 8.5), muted);
}
function cardsGrid(slide, items, x, y, w, h, cols = 3, opts = {}) {
  const gap = opts.gap || 0.16;
  const rows = Math.ceil(items.length / cols);
  const cw = (w - gap * (cols - 1)) / cols;
  const ch = (h - gap * (rows - 1)) / rows;
  items.forEach((it, i) => {
    const cx = x + (i % cols) * (cw + gap);
    const cy = y + Math.floor(i / cols) * (ch + gap);
    card(slide, cx, cy, cw, ch, it.kicker, it.head, it.text, it);
  });
}
function kpis(slide, vals, x, y, w, h) {
  line(slide, x, y, x + w, y, C.ink, 0.8);
  const cw = w / vals.length;
  vals.forEach((v, i) => {
    meta(slide, v.kicker, x + i * cw, y + 0.16, cw - 0.1);
    add(slide, v.value, x + i * cw, y + 0.42, cw - 0.1, h - 0.45, {
      fontFace: F.en, fontSize: v.size || 27, color: v.accent ? C.accent : C.ink,
      lineSpacingMultiple: 0.78,
    });
    body(slide, v.text || '', x + i * cw, y + 0.84, cw - 0.1, 0.2, 7.2);
  });
}
function addImage(slide, imageName, x, y, w, h) {
  slide.addImage({ path: IMG(imageName), x, y, w, h });
  slide.addShape(pptx.ShapeType.rect, { x, y, w, h, fill: { color: C.paper, transparency: 100 }, line: { color: C.grey2, width: 0.4 } });
}
function twoCol(slide, left, right, num, metaText, titleText) {
  slideBase(slide, metaText, num);
  title(slide, titleText, 0.62, 1.05, 11.8, 0.9, 29);
  card(slide, 0.68, 2.3, 5.8, 3.7, left.kicker, left.head, left.text, left.opts || {});
  card(slide, 6.85, 2.3, 5.8, 3.7, right.kicker, right.head, right.text, right.opts || {});
}
function statement(slide, text, kicker, num, accent = false) {
  slideBase(slide, kicker, num, accent);
  title(slide, text, 0.68, 1.35, 11.8, 4.1, accent ? 46 : 44, accent ? C.white : C.ink);
  line(slide, 0.68, 6.25, 12.65, 6.25, accent ? C.white : C.ink, 0.8);
  meta(slide, kicker, 0.68, 6.42, 7, accent, accent ? 'EDEDE8' : C.text2);
}

const slides = [];
function make(num, metaText, notesExtra, fn) {
  const s = pptx.addSlide();
  fn(s);
  s.addNotes(noteFor(num, notesExtra));
  slides.push(s);
}

make(1, 'Title', '開場：用一個問題帶出專利、營業秘密與技術授權。', (s) => {
  slideBase(s, 'Trade Secrets IPR × Technology Licensing', 1);
  rect(s, 0, 0, W, H, C.paper);
  meta(s, '01 營業秘密   02 技術授權   03 執行優勢', 0.75, 1.0, 6);
  title(s, '營業秘密 ×\n技術授權', 0.72, 1.35, 7.2, 2.1, 45);
  add(s, 'Trade Secrets IPR — Technology Licensing', 0.78, 3.45, 6.6, 0.35, { fontFace: F.en, fontSize: 16, color: C.accent });
  body(s, '如果全世界都知道晶片怎麼設計、電視怎麼做聲音、自駕車怎麼看路，為什麼不是每家公司都做得一樣好？', 0.78, 4.35, 8.0, 0.72, 17, C.ink);
  rect(s, 9.6, 1.1, 1.8, 1.8, C.accent);
  rect(s, 10.4, 2.55, 2.1, 2.1, C.ink);
  line(s, 8.6, 5.5, 12.6, 5.5, C.accent, 1.4);
  meta(s, '18–20 MIN · 6 SPEAKERS', 0.78, 6.55, 4);
});

make(2, 'Thesis', '核心命題：專利與營業秘密不是強弱差別，而是交易結構不同。', (s) => {
  rect(s, 0, 0, W / 2, H, C.accent);
  rect(s, W / 2, 0, W / 2, H, C.paper);
  slideBase(s, 'Core Thesis', 2);
  meta(s, 'PUBLIC BOUNDARY', 0.72, 1.05, 3, true, 'DDE6FF');
  title(s, '專利是\n公開的\n邊界線', 0.72, 1.45, 5.2, 2.75, 39, C.white);
  body(s, '公開發明細節，換取 20 年排他權。', 0.78, 5.9, 4.8, 0.5, 13, 'EDEDE8');
  meta(s, 'PRIVATE OPERATING SYSTEM', 7.1, 1.05, 4);
  title(s, '營業秘密是\n私有的\n執行系統', 7.1, 1.45, 5.2, 2.75, 39, C.ink);
  body(s, '不公開配方、參數、失敗資料庫與量產 know-how；只要守得住，就沒有到期日。', 7.15, 5.55, 4.9, 0.7, 13);
});

make(3, 'Legal Definition', '法律三要件：秘密性、經濟價值、合理保密措施。', (s) => {
  slideBase(s, 'Section 01 · Legal Definition', 3);
  meta(s, '營業秘密三要件', 0.72, 1.0, 3);
  title(s, '不是公司說秘密，\n就算營業秘密', 0.72, 1.35, 6.4, 1.5, 34);
  body(s, '法律看的是資訊本身的秘密性、經濟價值，以及公司是否真的採取合理保密措施。', 7.05, 1.48, 4.9, 0.7, 14);
  cardsGrid(s, [
    { kicker: 'LAYER 01', head: '不為公眾所知', text: '不是一般人能輕易取得，也不是產業內可直接查知的資訊。' },
    { kicker: 'LAYER 02', head: '具有經濟價值', text: '正因為它被保密，才帶來成本、速度、品質或市場優勢。', accent: true },
    { kicker: 'LAYER 03', head: '合理保密措施', text: '有實際運作的控制機制；不是只有蓋一個「機密」章。', ink: true },
  ], 0.72, 3.35, 11.9, 2.15, 3);
});

make(4, 'Patent vs Trade Secret', '比較專利與營業秘密的交易方式、保護期限與風險。', (s) => {
  twoCol(s,
    { kicker: 'A / PATENT', head: '專利\n公開 / 但可排除', text: '任何人都能讀你的發明；但在有效期間內，沒有許可不能做。\n\n· 保護期限通常 20 年\n· 可擋住獨立發明\n· 侵權比對相對清楚' },
    { kicker: 'B / TRADE SECRET', head: '營業秘密\n不公開 / 但怕外洩', text: '靠控制與保密維持優勢；一旦公開，價值可能瞬間崩塌。\n\n· 只要不洩漏，可無限期\n· 擋不住合法逆向工程\n· 舉證依賴日誌與內控', accent: true },
    4, 'Section 01 · Patent vs Trade Secret', '兩種保護，兩種交易');
});

make(5, 'Decision Matrix', '決策矩陣：公開後會怎樣，決定要專利還是保密。', (s) => {
  slideBase(s, 'Decision Matrix', 5);
  title(s, '真正問題是：\n公開之後會怎樣？', 0.72, 1.0, 6.4, 1.55, 36);
  cardsGrid(s, [
    { kicker: '01 / REVERSE', head: '逆向工程', text: '產品賣出去就能拆解，偏向專利；製程藏在工廠裡，偏向秘密。', accent: true },
    { kicker: '02 / LIFECYCLE', head: '生命週期', text: '20 年內淘汰可專利；價值可能超過 20 年，保密更有利。' },
    { kicker: '03 / DETECT', head: '侵權偵測', text: '產品上看得見可比對；競爭對手工廠內濫用很難查。' },
    { kicker: '04 / INVENT', head: '獨立發明', text: '別人很快會想到時靠專利；需要昂貴試錯才知道時靠秘密。', ink: true },
  ], 0.72, 3.0, 11.9, 2.85, 4);
});

make(6, 'Monetization', '營業秘密變現的三條路：溢價、授權抽成、風險降低。', (s) => {
  slideBase(s, 'Section 02 · Monetization', 6);
  meta(s, 'Know-how 怎麼變成錢？', 0.72, 0.95, 4);
  title(s, '營業秘密不是只放在保險箱裡，\n它可以被轉換成現金流、溢價與風險降低。', 0.72, 1.27, 11.5, 1.15, 29);
  card(s, 0.72, 3.0, 2.35, 2.55, 'THREE PATHS', '價值\n轉化\n路徑', '', { ink: true, headSize: 23 });
  cardsGrid(s, [
    { kicker: '01 / PREMIUM', head: '溢價', text: '良率更高、成本更低、品質更穩，因此同樣產品可以有更好毛利。' },
    { kicker: '02 / LICENSING', head: '授權抽成', text: '把 know-how、軟體、品牌、認證與支援打包成授權組合。', accent: true },
    { kicker: '03 / RISK REDUCTION', head: '風險降低', text: '防止技術流失、避免追趕者跳過研發學費。' },
  ], 3.35, 3.0, 9.0, 2.55, 3);
});

make(7, 'Inspection Paradox', '驗貨悖論：買家要看秘密才願意買，但看了就不再秘密。', (s) => {
  statement(s, '買家要先看到秘密\n才知道值不值得買；\n但看了秘密，\n就不是秘密了。', 'Inspection Paradox · 驗貨悖論', 7);
});

make(8, 'Tiered Disclosure', '分階段揭露：先賣信心，再逐步賣秘密。', (s) => {
  slideBase(s, 'Tiered Disclosure', 8);
  meta(s, '先賣信心，再賣秘密', 0.72, 0.95, 4);
  title(s, '把授權設計成分階段產品，\n每一階段只揭露剛好足夠的資訊。', 0.72, 1.28, 10.9, 1.1, 28);
  const steps = [
    ['01', 'NDA + 評估包', '效果證明，不給配方'],
    ['02', '付費試用', '有限範圍與試產'],
    ['03', '簽約授權', 'SOP / 設備 / 材料'],
    ['04', '驗收里程碑', '重現後交進階參數'],
    ['05', '商業量產', '缺陷庫與下一代改進'],
  ];
  const x0 = 0.72, y = 3.2, gap = 0.12, cw = (11.9 - gap * 4) / 5;
  steps.forEach((st, i) => card(s, x0 + i * (cw + gap), y, cw, 2.3, st[0], st[1], st[2], { accent: i === 3 }));
  meta(s, 'DISCLOSURE IS A PRODUCT, NOT A LEAK.', 0.75, 6.2, 6);
});

make(9, 'Revenue Stack', '付款結構：前期費、里程碑、抽成、最低授權金、支援費、稽核權。', (s) => {
  slideBase(s, 'Revenue Stack', 9);
  meta(s, '付款結構就是風險分配', 0.72, 0.95, 4);
  title(s, '授權金不是一個數字，\n而是一整套揭露、驗收、使用與稽核機制。', 0.72, 1.25, 11.5, 1.15, 29);
  cardsGrid(s, [
    { kicker: '01', head: '前期授權金', text: '取得機密資料包的權利，也回收研發成本。' },
    { kicker: '02', head: '里程碑付款', text: '被授權方重現效能後，才進一步揭露。' },
    { kicker: '03', head: '按量／營收抽成', text: '分享商業化後的上行利潤。', accent: true },
    { kicker: '04', head: '最低授權金', text: '防止拿到技術後擱置不用。' },
    { kicker: '05', head: '技術支援費', text: '工程師培訓、現場調試、問題排除。' },
    { kicker: '06', head: '稽核權', text: '確認範圍、數量、地區與保密遵守。', ink: true },
  ], 0.72, 3.0, 11.9, 2.95, 3);
});

make(10, 'BMC', '商業模式畫布：營業秘密滲透關鍵資源、收入、成本、客戶關係。', (s) => {
  slideBase(s, 'BMC · Business Model Canvas', 10);
  title(s, '營業秘密滲透整個 BMC', 0.72, 0.95, 6.8, 0.55, 31);
  body(s, '它不只是關鍵資源；它也決定價值主張、收入來源、客戶關係與成本結構。', 0.75, 1.65, 6.7, 0.5, 13);
  const items = ['關鍵資源|製程配方、隱性 know-how', '價值主張|更快投產、更高良率', '收入來源|授權金、支援費、認證費', '關鍵活動|保護、改良、培訓、稽核', '客戶關係|長期支援、改進共享', '成本結構|安全、法律、監控、訴訟', '關鍵夥伴|被授權方、OEM、供應商', '通路|技轉包、受控資料室', '目標客群|缺時間、經驗與設備調校'].map((v, i) => { const [h, t] = v.split('|'); return { kicker: `BMC ${String(i + 1).padStart(2, '0')}`, head: h, text: t, accent: i === 2 }; });
  cardsGrid(s, items, 0.72, 2.25, 11.9, 3.75, 3, { gap: 0.12 });
  rect(s, 0.72, 6.25, 11.9, 0.44, C.ink);
  add(s, '專利變現的是排除權；營業秘密變現的是讓你更快、更便宜、更少損失地做到。', 0.9, 6.34, 11.3, 0.22, { fontSize: 10, color: C.white });
});

function caseIntro(s, cfg) {
  slideBase(s, cfg.meta, cfg.num);
  meta(s, cfg.kicker, 0.72, 1.0, 5);
  title(s, cfg.title, 0.72, 1.28, 5.6, 1.25, 34);
  addImage(s, cfg.image, 7.05, 0.95, 5.15, 1.18);
  body(s, cfg.lead, 7.05, 2.28, 5.2, 0.55, 13);
  cardsGrid(s, cfg.cards, 0.72, 3.05, 11.9, 2.15, 3, { gap: 0.14 });
  kpis(s, cfg.kpis, 0.72, 5.65, 11.9, 0.95);
}

make(11, 'Dolby Intro', 'Dolby 案：把 know-how、專利、軟體、品牌和認證打包成授權機器。', (s) => caseIntro(s, {
  num: 11, meta: 'Case 01 · Dolby', kicker: 'Dolby 授權生態系', title: '最乾淨的\n授權機器', image: '11-dolby.png',
  lead: 'Dolby 把專利、know-how、軟體、品牌與認證包成一個可測試、可稽核的授權系統。',
  cards: [
    { kicker: '01 / 晶片商', head: '實作授權', text: '技術整合進晶片' },
    { kicker: '02 / 設備商', head: '系統授權', text: '每台出貨付費', accent: true },
    { kicker: '03 / 內容平台', head: '格式輸出', text: '串流與標準綁定' },
    { kicker: '04 / 消費者', head: '品質訊號', text: '品牌成為期待' },
    { kicker: '05 / 認證', head: '上市前測試', text: '控制體驗與信任' },
    { kicker: '06 / 稽核', head: '收入防漏', text: '雙邊報告交叉驗證', ink: true },
  ],
  kpis: [
    { kicker: '授權收入', value: '$1.248B', text: 'FY2025' },
    { kicker: '營收占比', value: '92.5%', text: '來自授權', accent: true },
    { kicker: '被授權方', value: '~1,000', text: 'Licensees' },
  ],
}));

make(12, 'Dolby IP Bundle', 'Dolby 賣的不是單一專利，而是可測試、可認證、可稽核的 IP 組合。', (s) => {
  slideBase(s, 'Dolby · IP Bundle', 12);
  meta(s, '授權包不只是專利許可', 0.72, 0.95, 4);
  title(s, 'Dolby 賣的是一個可測試、可認證、可稽核的 IP 組合。', 0.72, 1.25, 11.2, 0.92, 30);
  cardsGrid(s, [
    { kicker: '01', head: '專利', text: '音訊壓縮、Dolby Atmos、Dolby Vision HDR。' },
    { kicker: '02', head: 'Know-how', text: '營業秘密、設計、規格、專有資訊。', accent: true },
    { kicker: '03', head: 'Show-how', text: '教學、技能傳授、測試結果與技術示範。' },
    { kicker: '04', head: '軟體', text: '參考解碼器，禁止逆向工程與修改。' },
    { kicker: '05', head: '品牌', text: 'Dolby 名稱與雙 D 標誌成為品質訊號。' },
    { kicker: '06', head: '認證', text: '產品上市前必須通過 Dolby 測試。', ink: true },
  ], 0.72, 2.75, 11.9, 3.1, 3);
});

make(13, 'Dolby Two-tier', '雙層授權：晶片商可整合技術，但只能賣給 Dolby 認可系統授權方。', (s) => {
  slideBase(s, 'Dolby · Two-tier Licensing', 13);
  meta(s, '雙層制確保每台設備都有付錢', 0.72, 0.95, 4);
  title(s, '晶片商可以整合 Dolby 技術，\n但只能賣給 Dolby 認可的系統授權方。', 0.72, 1.25, 11.5, 1.15, 29);
  cardsGrid(s, [
    { kicker: '01 / Implementation', head: '晶片商', text: 'Qualcomm、MediaTek 等把技術做進晶片，但銷售對象受控。' },
    { kicker: '02 / System', head: '設備商', text: 'Samsung、LG、Apple 等終端產品，每台出貨付授權金。', accent: true },
    { kicker: '03 / Certified', head: '合格名單', text: '通過測試後列入 SQR，才能取得含 Dolby 技術的晶片。' },
  ], 0.72, 3.15, 11.9, 1.75, 3);
  line(s, 1.5, 5.7, 11.5, 5.7, C.ink, 0.9);
  add(s, 'CHIP → DEVICE → LICENSE', 0.72, 5.92, 6.2, 0.45, { fontFace: F.en, fontSize: 25, color: C.ink });
  body(s, '控制供應鏈入口，就是控制授權收入出口。', 8.4, 5.95, 3.7, 0.35, 12);
});

make(14, 'Dolby Audit Loop', '交叉驗證：晶片商報告與設備商報告互相校驗，少報超過 5% 觸發三倍未付授權金。', (s) => {
  slideBase(s, 'Dolby · Audit Loop', 14);
  meta(s, '稽核不是防禦，是營收保護', 0.72, 0.95, 4);
  title(s, 'Dolby 同時拿晶片供應商與系統授權方的季度報告，\n用交叉驗證找出少報。', 0.72, 1.25, 11.7, 1.15, 27);
  const steps = ['晶片商回報出貨給誰', '設備商回報產品出貨', '兩組數據交叉比對', '少報超過 5% → 三倍未付授權金'];
  steps.forEach((t, i) => { add(s, `0${i + 1}`, 0.72, 3.0 + i * 0.55, 0.4, 0.25, { fontFace: F.en, fontSize: 16, color: i === 3 ? C.accent : C.ink }); body(s, t, 1.22, 3.02 + i * 0.55, 4.4, 0.28, 12, C.ink); line(s, 0.72, 2.92 + i * 0.55, 5.55, 2.92 + i * 0.55, C.grey2, 0.6); });
  card(s, 6.15, 2.8, 2.0, 0.95, 'SOURCE A', '晶片報告', '晶片商回報出貨對象');
  card(s, 10.0, 2.8, 2.0, 0.95, 'SOURCE B', '設備報告', '設備商回報產品出貨', { accent: true });
  card(s, 7.85, 4.1, 2.4, 1.0, 'CROSS-VERIFY', '稽核', '雙邊報告交叉驗證', { ink: true, headSize: 21 });
  line(s, 8.15, 3.75, 8.05, 4.1, C.ink, 1.1, true);
  line(s, 10.0, 3.75, 9.65, 4.1, C.accent, 1.1, true);
  line(s, 9.05, 5.1, 9.05, 5.55, C.ink, 1.1, true);
  rect(s, 7.0, 5.65, 4.25, 0.75, C.grey);
  meta(s, 'MISMATCH DETECTED', 7.2, 5.78, 2.3, false, C.accent);
  add(s, '少報超過 5% → 三倍未付授權金', 7.2, 6.02, 3.9, 0.25, { fontSize: 15, color: C.ink });
});

make(15, 'TSMC Intro', '台積電案：最強模式是不授權，把秘密留在工廠裡賣結果。', (s) => caseIntro(s, {
  num: 15, meta: 'Case 02 · TSMC', kicker: '台積電製程 know-how 系統', title: '最強的模式\n是不授權', image: '15-tsmc.png',
  lead: '台積電把秘密留在工廠裡，賣給客戶的是更高良率、更快投產與更可靠量產。',
  cards: [
    { kicker: '01 / 製程窗口', head: '溫度壓力', text: '氣體、時序、雜質容忍' },
    { kicker: '02 / 設備配方', head: '機台調校', text: '每台機器的最佳參數' },
    { kicker: '03 / 良率資料', head: '失敗記憶', text: '知道什麼不能做', accent: true },
    { kicker: '04 / 缺陷診斷', head: '早期訊號', text: '預測良率崩潰' },
    { kicker: '05 / 供應商', head: '材料批次', text: '化學品與變異控制' },
    { kicker: '06 / 客戶協作', head: '量產節奏', text: '投產、迭代、保密', ink: true },
  ],
  kpis: [
    { kicker: '2024 營收', value: 'US$90.08B', text: '約 NT$2.894T', size: 25 },
    { kicker: '先進製程', value: '69%', text: '7nm 以下營收占比' },
    { kicker: '內部登錄', value: '475,462', text: '2013–2024 營業秘密', accent: true, size: 27 },
  ],
}));

make(16, 'No-license', '台積電策略：授權出去等於失去秘密並創造競爭對手。', (s) => {
  rect(s, 0, 0, W / 2, H, C.accent);
  rect(s, W / 2, 0, W / 2, H, C.paper);
  slideBase(s, 'No-license Strategy', 16);
  title(s, '把秘密\n留著', 0.75, 1.5, 4.8, 1.7, 43, C.white);
  body(s, '授權出去等於同時失去秘密，並創造競爭對手。', 0.8, 5.75, 4.7, 0.45, 13, 'EDEDE8');
  title(s, '賣的是\n我幫你做', 7.15, 1.5, 4.9, 1.7, 43, C.ink);
  body(s, '客戶買到的不是配方，而是台積電把配方轉成量產結果的能力。', 7.2, 5.55, 4.8, 0.55, 13);
});

make(17, 'TSMC System', '製程 know-how 是參數、設備、良率、缺陷、供應商與負面知識的耦合系統。', (s) => {
  slideBase(s, 'TSMC · Know-how System', 17);
  meta(s, '台積電的 know-how 不是一個配方', 0.72, 0.95, 5);
  title(s, '它是一整套製造系統：參數、設備、良率、缺陷、\n供應商與負面知識彼此耦合。', 0.72, 1.25, 11.6, 1.1, 28);
  cardsGrid(s, [
    { kicker: 'CORE', head: '製程窗口', text: '每一步的溫度、壓力、氣體、時間。', accent: true },
    { kicker: 'MIDDLE', head: '良率學習資料庫', text: '知道哪些條件會讓良率崩潰。' },
    { kicker: 'OUTER', head: '設備與供應鏈', text: '機台配方、材料批次、缺陷診斷與客戶協作。' },
  ], 0.72, 3.05, 11.9, 1.75, 3);
  line(s, 0.72, 5.65, 12.62, 5.65, C.ink, 0.8);
  add(s, '負面知識：知道什麼不能做', 0.72, 5.9, 6.8, 0.42, { fontSize: 24, color: C.accent });
  body(s, '重建秘密不只是重新發明，還要重新經歷所有失敗。', 8.0, 5.95, 4.0, 0.35, 12);
});

make(18, 'TSMC Boundary', '授權策略也包括決定哪些東西絕對不能放出去。', (s) => {
  twoCol(s,
    { kicker: 'A / LICENSE OUT', head: '授權出去\n失去秘密 / 創造對手', text: '製程窗口、設備配方與良率資料庫一旦外流，競爭者可以跳過多年試錯。\n\n· 揭露後不可逆\n· 濫用難偵測\n· 合約救濟晚於損害', opts: { ink: true } },
    { kicker: 'B / SERVICE MODEL', head: '自己代工\n保留秘密 / 出售結果', text: 'Apple、Nvidia、AMD 付的是代工費，實際買的是台積電的秘密製程能力。\n\n· 更高良率\n· 更快投產\n· 更低量產風險', opts: { accent: true } },
    18, 'TSMC · Boundary of Licensing', '授權策略也包括決定哪些東西絕對不能放出去。');
});

make(19, 'Waymo Intro', 'Waymo/Uber 案：營業秘密竊盜是時間壓縮攻擊。', (s) => caseIntro(s, {
  num: 19, meta: 'Case 03 · Waymo v. Uber', kicker: 'Waymo / Uber', title: '時間壓縮\n攻擊', image: '19-waymo.png',
  lead: '營業秘密竊盜買的不是資訊本身，而是跳過數年研發、試錯與人才訓練的速度。',
  cards: [
    { kicker: '01 / 下載', head: '14,107 檔', text: '9.7GB 設計與規格', accent: true },
    { kicker: '02 / 搬移', head: 'USB 硬碟', text: '外接儲存轉出' },
    { kicker: '03 / 抹除', head: '重灌筆電', text: '試圖破壞證據' },
    { kicker: '04 / 意外', head: 'OTTO FILES', text: '供應商信件 CC' },
    { kicker: '05 / 鑑識', head: '三組日誌', text: 'Armada / Bit9 / GRR' },
    { kicker: '06 / 後果', head: '$245M', text: '和解與禁用承諾', ink: true },
  ],
  kpis: [
    { kicker: '下載檔案', value: '14,107', text: '共 9.7GB', accent: true },
    { kicker: '民事和解', value: '$245M', text: '0.34% Uber 股權' },
    { kicker: 'Uber 投入', value: '~$3B', text: '最終出售自駕部門' },
  ],
}));

make(20, 'Waymo Timeline', '時間線：下載、收購、意外 CC、和解、特赦。', (s) => {
  slideBase(s, 'Waymo · Timeline', 20);
  title(s, '一條偷竊、收購、曝光、和解的時間線', 0.72, 1.0, 10.8, 0.6, 29);
  body(s, '真正的戲劇性不在法庭，而在一封被意外 CC 的供應商信件。', 0.75, 1.72, 7.8, 0.4, 13);
  const steps = [
    ['2015/12', '下載與抹除', '14,107 檔案到 USB'],
    ['2016/08', 'Uber 收購 Otto', '約 $680M'],
    ['2016/12', '意外 CC', 'OTTO FILES 附件曝光'],
    ['2018/02', '開庭後和解', '0.34% 股權'],
    ['2021/01', '特赦', '未入獄服刑'],
  ];
  line(s, 0.9, 4.0, 12.2, 4.0, C.ink, 0.9);
  const gap = 11.3 / 4;
  steps.forEach((st, i) => {
    const x = 0.9 + i * gap;
    rect(s, x - 0.04, 3.94, 0.08, 0.08, i === 2 ? C.accent : C.ink);
    meta(s, st[0], x - 0.35, i % 2 ? 4.35 : 2.95, 1.0, false, i === 2 ? C.accent : C.text2);
    add(s, st[1], x - 0.65, i % 2 ? 4.65 : 3.22, 1.5, 0.25, { fontSize: 11, color: i === 2 ? C.accent : C.ink, align: 'center' });
    body(s, st[2], x - 0.7, i % 2 ? 4.95 : 3.52, 1.6, 0.3, 8.5, C.text2);
  });
  meta(s, 'THE SMOKING GUN WAS AN EMAIL.', 0.75, 6.35, 5);
});

make(21, 'Waymo Forensics', '鑑識：筆電被擦掉，但日誌已回傳伺服器。', (s) => {
  slideBase(s, 'Waymo · Forensics', 21);
  meta(s, '筆電被擦掉，日誌還在', 0.72, 0.95, 4);
  title(s, 'Google 的安全工具在筆電被銷毀前，\n已經把關鍵存取紀錄傳回伺服器。', 0.72, 1.25, 11.6, 1.1, 28);
  card(s, 0.72, 3.0, 2.35, 2.45, 'EVIDENCE STACK', '三組\n鑑識\n工具', '', { ink: true, headSize: 23 });
  cardsGrid(s, [
    { kicker: '01', head: 'Armada', text: '確認下載來源來自 Levandowski 的公司筆電。' },
    { kicker: '02', head: 'Bit9', text: '記錄 USB 外接硬碟的連接與移除。', accent: true },
    { kicker: '03', head: 'GRR', text: '在筆電被擦掉前，把存取日誌回傳伺服器。' },
  ], 3.35, 3.0, 9.0, 2.45, 3);
});

make(22, 'Waymo Outcome', '結果：紅旗、刑案、刑期、Uber 自駕投入歸零。', (s) => {
  slideBase(s, 'Waymo · Red Flags / Outcome', 22);
  const rows = [
    ['50,000', '封工作信件在個人電腦上', 'Stroz 盡調紅旗', true],
    ['33', '項竊盜罪起訴', '認罪 1 項'],
    ['18', '個月刑期', '後被特赦，未服刑'],
    ['$3B', 'Uber 自駕累計投入', '最終出售 ATG'],
  ];
  rows.forEach((r, i) => {
    const y = 1.0 + i * 1.28;
    line(s, 0.72, y + 1.08, 12.4, y + 1.08, C.grey2, 0.7);
    add(s, r[0], 0.72, y, 3.0, 0.8, { fontFace: F.en, fontSize: 42, color: r[3] ? C.accent : C.ink, lineSpacingMultiple: 0.75 });
    add(s, r[1], 4.0, y + 0.15, 5.8, 0.38, { fontSize: 18, color: C.ink });
    meta(s, r[2], 4.0, y + 0.7, 3.5);
    rect(s, 11.4, y + 0.24, 0.55, 0.55, r[3] ? C.accent : C.grey2);
  });
});

make(23, 'Defense Architecture', '合理保密措施：法律、員工、技術、組織、證據、實體與供應鏈。', (s) => {
  slideBase(s, 'Defense Architecture', 23);
  meta(s, '內控不只是防禦，它是變現的基礎建設', 0.72, 0.95, 5);
  title(s, '如果你不能證明合理保密措施，\n法律上你的秘密可能根本不算營業秘密。', 0.72, 1.25, 11.6, 1.1, 28);
  cardsGrid(s, [
    { kicker: '01', head: '法律層', text: 'NDA、發明歸屬、保密存續、稽核權。' },
    { kicker: '02', head: '員工層', text: '角色式存取、培訓紀錄、離職面談。' },
    { kicker: '03', head: '技術層', text: 'Need-to-know、DLP、浮水印、日誌、禁止批量匯出。', accent: true },
    { kicker: '04', head: '組織層', text: '營業秘密登錄冊、揭露審批、供應商管控。' },
    { kicker: '05', head: '證據層', text: '不可竄改日誌、版本控管、存取歷史。' },
    { kicker: '06', head: '實體／供應鏈', text: '無手機區、遠端工作規則、現場稽核。', ink: true },
  ], 0.72, 3.0, 11.9, 2.95, 3);
});

make(24, 'Closing', '收束：營業秘密是可控的重複性；控制好就是現金流，控制不好價值崩塌。', (s) => {
  rect(s, 0, 0, W / 2, H, C.accent);
  rect(s, W / 2, 0, W / 2, H, C.paper);
  slideBase(s, 'Closing', 24);
  meta(s, 'MANIFESTO', 0.72, 1.0, 2.5, true, 'DDE6FF');
  title(s, '秘密不只是資訊。\n它是可控的\n重複性。', 0.72, 1.35, 5.5, 2.7, 36, C.white);
  meta(s, 'END · TRADE SECRETS IPR', 0.72, 6.35, 4.5, true, 'DDE6FF');
  meta(s, 'THREE MODES', 7.05, 1.0, 3);
  const take = [
    ['01', 'Dolby', '把 know-how 打包進授權組合，每年約 $1.25B 授權收入。'],
    ['02', '台積電', '把 know-how 留著，賣服務與量產能力。'],
    ['03', 'Waymo / Uber', 'know-how 被偷走，和解 $245M，Uber 最終放棄自駕。'],
  ];
  take.forEach((t, i) => {
    const y = 1.55 + i * 1.2;
    line(s, 7.05, y - 0.12, 12.2, y - 0.12, i === 2 ? C.accent : C.grey2, i === 2 ? 1.4 : 0.6);
    add(s, t[0], 7.05, y, 0.42, 0.35, { fontFace: F.en, fontSize: 22, color: i === 2 ? C.accent : C.ink });
    add(s, t[1], 7.75, y + 0.02, 2.0, 0.28, { fontSize: 14, color: i === 2 ? C.accent : C.ink });
    body(s, t[2], 7.75, y + 0.4, 4.2, 0.45, 10);
  });
  body(s, '控制好了，就是現金流；控制不好，價值一夕崩塌。', 7.05, 6.32, 5.2, 0.3, 11, C.text2);
});

fs.mkdirSync(path.dirname(OUT), { recursive: true });
pptx.writeFile({ fileName: OUT });
console.log(`Wrote ${OUT}`);
