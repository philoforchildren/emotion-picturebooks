const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
        AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign } = require('docx');

const INK = '4A3B2E', INK2 = '7A6A58', CREAM = 'F3EDE2', LINE = 'D8CDBC';

// 19 本精選:主題 → [書名, 適讀年齡, 一句話推薦]
const GROUPS = [
['◆ 認識情緒', [
 ['情緒都在忙什麼？','低～高年級','把每種情緒畫成可愛角色，看見它們怎麼影響我們'],
 ['情緒大飯店','學齡前～中年級','每種情緒都是來住宿的房客，都值得被好好接待'],
]],
['◆ 生氣・不耐煩', [
 ['問個不停的小孩，加斯東2','學齡前～中年級','為什麼愛生氣？跟孩子一起想一想的親子哲學繪本'],
 ['大吼吼','低～中年級','等太久好煩！幽默化解失去耐性的怒氣'],
]],
['◆ 害怕・恐懼', [
 ['害怕沒關係 大人也會怕喔！','學齡前～低年級','爸爸坦白說：大人也會怕，我們一起面對'],
 ['門外有一頭獅子','學齡前～高年級','用黃色與藍色，看見恐懼與勇氣的心境變化'],
]],
['◆ 焦慮・煩惱', [
 ['我有一些小煩惱','學齡前～低年級','把說不出口的煩惱畫出來，學著和它們相處'],
 ['賽米希的煩惱','學齡前～中年級','煩惱像石頭壓在心上——問問大家都怎麼辦？'],
]],
['◆ 悲傷・失落・道別', [
 ['悲傷，讓我抱抱你','學齡前～中年級','悲傷來敲門時別害怕，它只想知道自己不孤單'],
 ['和阿布說再見','學齡前～低年級','心愛的寵物離開後，陪孩子慢慢走過想念'],
 ['咦，帽子不見了？','學齡前～中年級','東西不見是孩子的大失落，用幽默溫柔接住'],
]],
['◆ 嫉妒・羨慕・比較', [
 ['月亮的新衣','學齡前～高年級','用月亮的圓缺，陪孩子認識羨慕和嫉妒'],
 ['蹦蹦和跳跳','學齡前～高年級','朋友比我厲害，我們還能當好朋友嗎？'],
]],
['◆ 孤單・寂寞', [
 ['想要被擁抱的仙人掌','學齡前～中年級','滿身刺的仙人掌，也好想被抱一下'],
]],
['◆ 害羞・內向', [
 ['害羞的小鼴鼠','學齡前～中年級','擔心個不停的小鼴鼠，鼓起勇氣交到新朋友'],
]],
['◆ 表達與調節', [
 ['熱狗','低～中年級','凱迪克金獎：又煩又熱的心，被海闊天空接住'],
 ['想哭就哭成一座噴水池','學齡前～低年級','想哭就哭吧！眼淚也可以變成好玩的想像'],
 ['我的眼淚果醬','學齡前～中年級','孤單、委屈、開心，都可以煮進果醬裡'],
]],
['◆ 家庭裡的情緒', [
 ['爸爸的小黑雲','低～高年級','大人也有憂鬱的時候，孩子的關心是穿透烏雲的光'],
]],
];

const PAGE_W = 11906, MARGIN = 720, CW = PAGE_W - MARGIN*2; // A4 內容寬 10466
const COLS = [3460, 2300, 4706]; // 書名 | 適讀 | 推薦 = 10466
const bd = { style: BorderStyle.SINGLE, size: 2, color: LINE };
const borders = { top: bd, bottom: bd, left: bd, right: bd };
const cellMargin = { top: 26, bottom: 26, left: 110, right: 110 };

function cell(text, {w, bold=false, fill=null, color=INK, size=18, align=AlignmentType.LEFT, span=1}={}){
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA }, margins: cellMargin,
    columnSpan: span, verticalAlign: VerticalAlign.CENTER,
    shading: fill ? { fill, type: ShadingType.CLEAR } : undefined,
    children: [new Paragraph({ alignment: align,
      children: [new TextRun({ text, bold, color, size })] })],
  });
}

const rows = [
  new TableRow({ tableHeader: true, children: [
    cell('書名', {w:COLS[0], bold:true, fill:INK, color:'FFFFFF', size:20}),
    cell('適讀年齡', {w:COLS[1], bold:true, fill:INK, color:'FFFFFF', size:20, align:AlignmentType.CENTER}),
    cell('這本書在談什麼', {w:COLS[2], bold:true, fill:INK, color:'FFFFFF', size:20}),
  ]}),
];
for (const [theme, books] of GROUPS){
  rows.push(new TableRow({ children: [
    cell(theme, {w:CW, bold:true, fill:CREAM, size:20, span:3}),
  ]}));
  for (const [name, age, note] of books){
    rows.push(new TableRow({ children: [
      cell('《' + name + '》', {w:COLS[0], bold:true}),
      cell(age, {w:COLS[1], align:AlignmentType.CENTER, color:INK2}),
      cell(note, {w:COLS[2], color:INK2}),
    ]}));
  }
}

const qr = fs.readFileSync(String.raw`D:\Google Drive串流\我的雲端硬碟\Amy-Agent\100_Todo\drafts\storytelling\2026-06-10_情緒繪本書單QRcode.png`);

// 底部:左說明 + 右 QR 的兩欄無框表格
const noBd = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: noBd, bottom: noBd, left: noBd, right: noBd };
const footer = new Table({
  width: { size: CW, type: WidthType.DXA }, columnWidths: [8866, 1600],
  rows: [ new TableRow({ children: [
    new TableCell({ borders: noBorders, width:{size:8866,type:WidthType.DXA},
      verticalAlign: VerticalAlign.CENTER, margins: {top:60,bottom:0,left:0,right:160},
      children: [
        new Paragraph({ spacing:{after:60}, children:[ new TextRun({ text:'掃描 QR code 看完整 80 本情緒繪本書單', bold:true, size:20, color:INK }) ]}),
        new Paragraph({ spacing:{after:60}, children:[ new TextRun({ text:'可依 10 大情緒主題、適讀年齡篩選，每本都有書封、簡介與評審推薦，手機就能看。', size:17, color:INK2 }) ]}),
        new Paragraph({ children:[ new TextRun({ text:'書目選自文化部「中小學生讀物選介」第 42–47 次獲選書單（2020–2025），各主題均可在圖書館借閱。', size:15, color:INK2 }) ]}),
      ]}),
    new TableCell({ borders: noBorders, width:{size:1600,type:WidthType.DXA},
      verticalAlign: VerticalAlign.CENTER, margins:{top:60,bottom:0,left:0,right:0},
      children: [ new Paragraph({ alignment: AlignmentType.RIGHT, children: [
        new ImageRun({ type:'png', data: qr, transformation:{ width:95, height:95 },
          altText:{ title:'書單 QR code', description:'情緒繪本書單網站 QR code', name:'qrcode' } }),
      ]})]}),
  ]})]
});

const doc = new Document({
  styles: { default: { document: { run: { font: '微軟正黑體', size: 18 }, paragraph: { spacing: { line: 238 } } } } },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN } } },
    children: [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing:{after:40},
        children: [new TextRun({ text:'暑假親子共讀・情緒繪本精選', bold:true, size:32, color:INK })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing:{after:100},
        children: [new TextRun({ text:'從文化部歷屆「中小學生讀物選介」獲選好書中，為孩子挑出 19 本談情緒的繪本', size:19, color:INK2 })] }),
      new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: COLS, rows }),
      new Paragraph({ spacing:{before:70}, children:[] }),
      footer,
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  const out = String.raw`D:\Google Drive串流\我的雲端硬碟\Amy-Agent\100_Todo\drafts\storytelling\2026-06-10_暑假親子共讀情緒繪本精選.docx`;
  fs.writeFileSync(out, buf);
  console.log('完成:', out);
});
