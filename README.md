const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, LevelFormat, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageBreak
} = require('docx');
const fs = require('fs');
 
const BD = "1F4E79", BM = "2E75B6", BL = "BDD7EE", BP = "DEEAF1";
const OR = "C55A11", OL = "FCE4D6";
const GD = "375623", GL = "E2EFDA";
const GR = "F2F2F2", WH = "FFFFFF";
const RD = "C00000", RL = "FFE7E7";
 
const b1 = { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" };
const bs = { top: b1, bottom: b1, left: b1, right: b1 };
const bn = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const bns = { top: bn, bottom: bn, left: bn, right: bn };
 
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, font: "Arial", size: 34, bold: true, color: WH })],
    shading: { fill: BD, type: ShadingType.CLEAR },
    spacing: { before: 320, after: 100 },
    indent: { left: 180 },
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: "Arial", size: 26, bold: true, color: WH })],
    shading: { fill: BM, type: ShadingType.CLEAR },
    spacing: { before: 240, after: 80 },
    indent: { left: 140 },
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, font: "Arial", size: 23, bold: true, color: BD })],
    spacing: { before: 180, after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 3, color: BL } }
  });
}
function body(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({ text, font: "Arial", size: 21, bold: opts.bold || false })],
    spacing: { before: 60, after: 60 },
    indent: opts.indent ? { left: 360 } : {}
  });
}
function space() { return new Paragraph({ children: [new TextRun(" ")], spacing: { before: 30, after: 30 } }); }
 
function tipBox(text) {
  return new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [9026], rows: [
    new TableRow({ children: [new TableCell({
      borders: { top: { style: BorderStyle.SINGLE, size: 4, color: OR }, bottom: bn, left: { style: BorderStyle.SINGLE, size: 12, color: OR }, right: bn },
      shading: { fill: OL, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 160, right: 160 },
      width: { size: 9026, type: WidthType.DXA },
      children: [new Paragraph({ children: [
        new TextRun({ text: "💡 ポイント　", font: "Arial", size: 21, bold: true, color: OR }),
        new TextRun({ text, font: "Arial", size: 21 })
      ], spacing: { before: 40, after: 40 } })]
    })] })
  ]});
}
 
function warnBox(text) {
  return new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [9026], rows: [
    new TableRow({ children: [new TableCell({
      borders: { top: { style: BorderStyle.SINGLE, size: 4, color: RD }, bottom: bn, left: { style: BorderStyle.SINGLE, size: 12, color: RD }, right: bn },
      shading: { fill: RL, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 160, right: 160 },
      width: { size: 9026, type: WidthType.DXA },
      children: [new Paragraph({ children: [
        new TextRun({ text: "⚠️ 注意　", font: "Arial", size: 21, bold: true, color: RD }),
        new TextRun({ text, font: "Arial", size: 21 })
      ], spacing: { before: 40, after: 40 } })]
    })] })
  ]});
}
 
function voiceBox(title, text) {
  return new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [9026], rows: [
    new TableRow({ children: [new TableCell({
      borders: bs, shading: { fill: GL, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 160, right: 160 },
      width: { size: 9026, type: WidthType.DXA },
      children: [
        new Paragraph({ children: [new TextRun({ text: title, font: "Arial", size: 20, bold: true, color: GD })], spacing: { before: 40, after: 20 } }),
        new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 20, color: "333333", italics: true })], spacing: { before: 20, after: 40 } })
      ]
    })] })
  ]});
}
 
function blueBox(title, items) {
  return new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [9026], rows: [
    new TableRow({ children: [new TableCell({
      borders: bs, shading: { fill: BD, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 },
      width: { size: 9026, type: WidthType.DXA },
      children: [new Paragraph({ children: [new TextRun({ text: title, font: "Arial", size: 22, bold: true, color: WH })], spacing: { before: 40, after: 40 } })]
    })] }),
    ...items.map(item => new TableRow({ children: [new TableCell({
      borders: bs, shading: { fill: BP, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 160, right: 160 },
      width: { size: 9026, type: WidthType.DXA },
      children: [new Paragraph({ children: [
        new TextRun({ text: "▶  ", font: "Arial", size: 21, bold: true, color: BM }),
        new TextRun({ text: item, font: "Arial", size: 21 })
      ], spacing: { before: 40, after: 40 } })]
    })] }))
  ]});
}
 
function tbl(headers, rows, widths) {
  return new Table({
    width: { size: 9026, type: WidthType.DXA }, columnWidths: widths,
    rows: [
      new TableRow({ children: headers.map((h, i) => new TableCell({
        borders: bs, shading: { fill: BM, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        width: { size: widths[i], type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, font: "Arial", size: 19, bold: true, color: WH })] })]
      })) }),
      ...rows.map((row, ri) => new TableRow({ children: row.map((cell, i) => new TableCell({
        borders: bs, shading: { fill: ri % 2 === 0 ? WH : BP, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        width: { size: widths[i], type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: cell, font: "Arial", size: 19 })] })]
      })) }))
    ]
  });
}
 
function check(text) {
  return new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [500, 8526], rows: [
    new TableRow({ children: [
      new TableCell({ borders: bs, shading: { fill: WH, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 80, right: 80 }, width: { size: 500, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "□", font: "Arial", size: 22 })] })] }),
      new TableCell({ borders: bs, shading: { fill: GR, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, width: { size: 8526, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 21 })] })] })
    ]})
  ]});
}
 
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 34, bold: true, font: "Arial", color: WH }, paragraph: { spacing: { before: 320, after: 100 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: WH }, paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 23, bold: true, font: "Arial", color: BD }, paragraph: { spacing: { before: 180, after: 60 }, outlineLevel: 2 } },
    ]
  },
  numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } } },
    children: [
 
      // ===== 表紙 =====
      new Paragraph({ children: [new TextRun({ text: "SHOWA GROUP　営業新入社員 研修教材", font: "Arial", size: 19, color: "777777" })], spacing: { before: 0, after: 80 } }),
      new Paragraph({ children: [new TextRun({ text: "商品・テクノロジー 研修テキスト", font: "Arial", size: 50, bold: true, color: BD })], spacing: { before: 200, after: 60 } }),
      new Paragraph({ children: [new TextRun({ text: "SUGOIE ／ RAKUIE ／ 性能比較・換気・断熱・建物仕様の総合理解", font: "Arial", size: 23, color: BM })], spacing: { before: 60, after: 200 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BM } } }),
      space(),
      new Table({ width: { size: 9026, type: WidthType.DXA }, columnWidths: [4400, 4626], rows: [new TableRow({ children: [
        new TableCell({ borders: bns, shading: { fill: WH, type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 0, right: 0 }, width: { size: 4400, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "氏名：＿＿＿＿＿＿＿＿＿＿＿＿", font: "Arial", size: 21 })] })] }),
        new TableCell({ borders: bns, shading: { fill: WH, type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 0, right: 0 }, width: { size: 4626, type: WidthType.DXA }, children: [new Paragraph({ children: [new TextRun({ text: "研修日：＿＿＿＿年＿＿月＿＿日", font: "Arial", size: 21 })] })] }),
      ]}]}),
      space(),
 
      // ===== 1. この研修で身につけること =====
      h1("1. この研修で身につけること"),
      space(),
      body("この研修では、SUGOIE・RAKUIEの「商品知識」だけでなく、実際の営業現場で先輩が話している内容（換気・断熱・気候・商品使い分けなど）を体系的に学びます。"),
      space(),
      blueBox("この研修で身につける4つのこと", [
        "① 断熱材・換気システムの種類と特徴を自分の言葉で説明できる",
        "② 播磨エリアの気候特性を踏まえた「適切な性能提案」ができる",
        "③ SUGOIEとRAKUIEの仕様・価格・ターゲットの違いを即答できる",
        "④ お客様の予算やこだわりに合わせて商品・仕様を使い分けて提案できる",
      ]),
      space(),
      tipBox("「知識を覚える」のではなく「お客様に説明できる状態」を目指す。勉強会で先輩が話していた「引き出しをたくさん作って、すぐ出せる」状態がゴール。"),
      space(),
 
      // ===== 2. 播磨エリアの気候と性能提案の前提 =====
      h1("2. 播磨エリアの気候と性能提案の前提知識"),
      space(),
      h2("2-1　播磨エリアの気候特性"),
      space(),
      body("加古川・姫路・明石を含む「播磨エリア」は、日本の中でも比較的温暖で安定した気候のエリアです。これは性能提案をする上での大前提の知識です。"),
      space(),
      tbl(
        ["比較項目", "播磨エリア", "山陰・豊岡エリア", "北海道"],
        [
          ["冬の寒さ", "比較的穏やか", "積雪多く厳しい", "極寒・ガンガン雪"],
          ["夏の暑さ", "近年暑いが標準的", "同等〜やや涼しい", "短い夏"],
          ["太陽光発電", "日照時間多く発電量多い", "冬は曇りがちで発電少ない", "北向きは不利"],
          ["断熱の優先度", "高いほどよいが過剰になりにくい", "高断熱が重要", "北海道基準が必要"],
        ],
        [2200, 2000, 2200, 2626]
      ),
      space(),
      voiceBox("🎤 勉強会での先輩の言葉", "播磨地区はめちゃめちゃ安定している温暖な気候と言われてて。冬もそんなに冷え込むっていうこともないし、雪がガンガン降ることもない。だからえっと今性能は非常に言われているんだけど、まあ高値に越したことはないんやけど、じゃあ北海道基準の断熱はここで言いますかって言われたらオーバースペックな感じではある。まあインパクトや商品の特徴付けとしては響くから、そういうこともできるし。"),
      space(),
      tipBox("播磨エリアは気候に恵まれているため、「ZEH水準・断熱等級5でも十分すぎるくらいの性能」がある。北海道基準の断熱をアピールすることは「商品の特徴付け」として使えるが、予算が合わない場合は無理に推す必要はない。この前提を知っておくと、お客様への提案の幅が広がる。"),
      space(),
      h2("2-2　太陽光発電と気候の関係"),
      space(),
      body("太陽光発電は日照時間に左右されます。播磨エリアは晴れが多く、山陰地方などと比べて年間発電量が大きくなりやすいエリアです。"),
      space(),
      voiceBox("🎤 勉強会での先輩の言葉", "太陽光一つとしても、やっぱり山陰地方って発電量全然少ないんです冬なんかずっと曇ってるし。だからこの辺（播磨）は非常に恵まれた気候だよっていうところ。"),
      space(),
      tipBox("播磨エリアで太陽光発電を提案するときは「この辺は日照時間が多く、発電効率が高いエリアです」という一言を添えると説得力が増す。"),
      space(),
 
      // ===== 3. 断熱材の知識 =====
      h1("3. 断熱材の知識"),
      space(),
      h2("3-1　断熱材が重要な理由"),
      space(),
      body("家の断熱性能を高くするには断熱材が大事です。夏の暑さや冬の寒さをよせつけず、家中を快適な温度に保つには十分な断熱が必要です。断熱性が高ければ冷暖房コストも下がります。"),
      space(),
      h2("3-2　現場発泡断熱材（ウレタンフォーム）"),
      space(),
      body("SUGOIE・RAKUIEともに「現場発泡硬質ウレタンフォーム」を採用しています。これは現場で液状の材料を吹き付けて発泡させる断熱材です。"),
      space(),
      tbl(
        ["特徴", "内容・お客様へのメリット"],
        [
          ["高い省エネ性能", "隙間なく充填するため断熱ロスがなく、冷暖房費が削減できる"],
          ["高い気密性", "構造体に密着し複雑な形状にも隙間なく入るため、熱損失が大幅に減少"],
          ["遮音性能", "微細なセル構造が高い吸音性を発揮。外部の騒音をシャットアウト"],
          ["完全ノンフロン", "炭酸ガスを発泡材に使用。オゾン層破壊・地球温暖化に配慮した断熱材"],
          ["経年劣化が小さい", "連続気泡の内部は空気を内包するため、性能の初期値と経年劣化後の差が小さい"],
          ["カビ・ダニ防止", "微細連続気泡が湿気を吸放出し内部結露を抑制。カビ・ダニの繁殖を防ぐ"],
        ],
        [2800, 6226]
      ),
      space(),
      h3("施工厚みの比較（SUGOIE vs RAKUIE）"),
      space(),
      tbl(
        ["部位", "SUGOIE（すごい家）", "RAKUIE（らくいえ）"],
        [
          ["壁", "105mm（現場発泡）", "100mm（現場発泡ダルトフォーム）"],
          ["屋根", "210mm（現場発泡）", "210mm（現場発泡）"],
          ["床", "（パンフレット記載なし）", "105mm（サットフォーム）"],
        ],
        [2200, 3413, 3413]
      ),
      space(),
      voiceBox("🎤 勉強会での先輩の言葉", "こっちは吹き付けもなしかも厚みを変わらないですし、こちらが樹脂（窓）になってだいぶ違うよなで、アルゴンガスも入ってますし。断熱で言ったら「アルミサッシってイメージしてもらったら水がジャーってなってたり朝寒い日に結露がいっぱいついてたりとか、そういうのが起こりにくいんですよね」という話をしたり。"),
      space(),
      tipBox("断熱の説明はお客様がイメージしやすい「結露の話」で入ると伝わりやすい。「冬に窓に水滴がつくことありますよね。樹脂窓・断熱材の組み合わせでそれが起こりにくくなります。暖房代も変わります」と身近な体験に結びつけて話す。"),
      space(),
      h2("3-3　サッシ（窓）の断熱性能比較"),
      space(),
      tbl(
        ["項目", "SUGOIE（APW330 樹脂サッシ）", "RAKUIE（エピソードⅡNEO アルミ樹脂複合）"],
        [
          ["フレーム素材", "完全樹脂", "室内側：樹脂 / 室外側：アルミ"],
          ["ガラス", "Low-E複層ガラス（アルゴンガス入）", "Low-E複層ガラス（アルゴンガス入）"],
          ["熱貫流率（参考）", "1.31 W/(㎡・K)（APW330）", "アルミ樹脂複合で中程度"],
          ["断熱性能", "国内最高レベル", "一般的な高性能サッシ"],
          ["結露のしにくさ", "非常に高い", "樹脂部分は結露しにくい"],
        ],
        [2600, 3213, 3213]
      ),
      space(),
      tipBox("「樹脂か、アルミ樹脂複合か」はSUGOIEとRAKUIEの大きな違いの一つ。予算を抑えたいお客様にはRAKUIEの窓でも十分な性能であることを伝え、性能重視のお客様にはSUGOIEの完全樹脂窓のメリットを強調する。"),
      space(),
 
      // ===== 4. 換気システムの知識 =====
      h1("4. 換気システムの知識（重要）"),
      space(),
      h2("4-1　なぜ24時間換気が必要なのか"),
      space(),
      body("現代の高気密・高断熱住宅では、窓を閉めきることで室内に湿気がこもりやすくなります。湿気は家の大敵で、カビやダニの原因になるだけでなく、壁の中で結露が起きて構造体を腐らせます。"),
      space(),
      voiceBox("🎤 勉強会での先輩の言葉", "湿気は人間が今この空間におるだけでも、湿気がどんどん発生してくる。喋ってた息するだけでもどんどん湿気が出てくるし、当然お湯沸かしたり、ご飯作ったり、どんどん湿気が出てくる。でどっかで温度差が絶対やびやか外でキンキンに冷えたら血（水分）を水滴になってカビになる。だから湿気をいかに外に出すかっていうのが、家を長持ちさせる大前提の条件。"),
      space(),
      body("建築基準法により、住宅には1時間で家全体の空気の1/2を入れ替えられる24時間換気システムの設置が義務付けられています。"),
      space(),
      h2("4-2　換気の3種類を覚える"),
      space(),
      tbl(
        ["種類", "給気", "排気", "主な用途"],
        [
          ["第1種換気", "機械（ファン）", "機械（ファン）", "高性能住宅・熱交換型に使用"],
          ["第2種換気", "機械（ファン）", "自然", "クリーンルームなど特殊用途"],
          ["第3種換気", "自然（給気口）", "機械（ファン）", "一般住宅で最も多く使われる"],
        ],
        [2200, 2000, 2000, 2826]
      ),
      space(),
      voiceBox("🎤 勉強会での先輩の言葉", "第一種は機械から機械。第二は（給気が）機械から自然。第三は自然から機械。トイレがそれで、24時間換気いやこれなぜ必要？匂いもそうだけど、目的は匂いじゃない。湿気や家への大敵なんか湿気だから。"),
      space(),
      h2("4-3　SUGOIE採用：第1種熱交換型換気（ダクトレス・ケィ・マック）"),
      space(),
      body("第1種換気の特徴は「機械で給気・排気の両方を行う」ことに加え、「熱交換」ができることです。SUGOIEでは特に「ダクトレス熱交換型」を採用しています。"),
      space(),
      h3("熱交換型の仕組み"),
      space(),
      voiceBox("🎤 勉強会での先輩の言葉", "北海道でガンガンに暖炉で温めたあったかい空気が家の中にある。でも換気しないといけないから24時間で空気入れ替えましたって言ったら、あったかい空気が出て行きて、極寒の吹雪の空気が入ってきたわけよ。何してるようわからへんやんか。それもったいないでしょっていうことで、出ていく空気の熱を入ってきた冷たいキンキンの空気に移して、ちょっと弱めてから入れる。それが熱交換。"),
      space(),
      tbl(
        ["特徴", "内容"],
        [
          ["熱交換効率", "業界最高クラス 92%（フード未装着時）"],
          ["ダクトレス方式", "各部屋の壁にセラミック素子ユニットを設置。ダクト不要でカビリスクが低い"],
          ["セラミック素子の仕組み", "排気時に室内の熱をセラミックが蓄積。給気時に外気をそのセラミックで温めて取り入れる"],
          ["設置台数", "標準5台（各居室に配置）"],
          ["換気サイクル", "60〜80秒で給気・排気が切り替わる"],
          ["ランニングコスト", "電気代は月1,000円前後（継続稼働するため）"],
        ],
        [2600, 6426]
      ),
      space(),
      h3("ダクトあり第1種換気の問題点（なぜダクトレスを選ぶか）"),
      space(),
      voiceBox("🎤 勉強会での先輩の言葉", "一般的な第1種換気は家の天井などにダクト（パイプ）がいっぱいできる。そのダクトの中ってカビが生えやすいんやな。湿気が通ったり、暑い寒いの空気が通るところやから、冷やされて結露する。ここでカビが発生してしまったら一生そのカビ播きシステムみたいな感じになる。これが第1種換気のやっぱりネックで。うちのダクトレスはパイプを使わないのでカビが撒き散らされるシステムにならないということです。"),
      space(),
      h2("4-4　RAKUIE採用：第3種24時間換気"),
      space(),
      body("RAKUIEは最も一般的な「第3種換気（自然給気＋機械排気）」を採用しています。仕組みはシンプルで、トイレや廊下に換気扇を設置し、各部屋の給気口から自然に空気を取り込む方式です。"),
      space(),
      voiceBox("🎤 勉強会での先輩の言葉", "こっち（第3種）寄りのトークするんやったら、世の中で一番使われているのは第三種換気なんで。どう考えたってメンテナンスが楽やし、コストもかかりにくい。至ってシンプルなシステムやから。第3種換気、壊れたプロペラ回らんようになったっていうことだったら、まあ単純にプロペラが壊れただけやから。はいそれだけやから。本当にめちゃめちゃシンプルかつ低コストかつ低ランニングコスト。"),
      space(),
      tbl(
        ["比較項目", "第1種熱交換型（SUGOIE）", "第3種換気（RAKUIE）"],
        [
          ["コスト（本体）", "高め（2台で40万円前後の場合も）", "低い（扇風機レベルのシンプル構造）"],
          ["ランニングコスト", "月1,000円程度の電気代が継続", "ほぼゼロ（小さなモーターのみ）"],
          ["メンテナンス", "セラミックフィルター清掃が必要", "極めて簡単（プロペラ清掃のみ）"],
          ["熱交換", "あり（熱ロスを最小化）", "なし"],
          ["冬の光熱費効果", "高い（外気を暖めて取り込む）", "標準的"],
          ["カビリスク", "ダクトレスのため低い", "シンプルなため問題なし"],
          ["播磨エリアでの適性", "高性能重視・予算ある方に最適", "コスト重視・十分な性能"],
        ],
        [2800, 3113, 3113]
      ),
      space(),
      voiceBox("🎤 勉強会での先輩の言葉", "播磨地区において、まあ余裕があれば全然やっていただいた方がいいと思うんですというか。そうじゃないんやって予算も抑えたいよって言われたら無理にやられるシステムでもないかもしれないね。3種換気壊れたプロペラ回らんようになったってことだったら、まあ単純に扇風機と一緒。電気的な問題よ。だから壊れ方もシンプル。"),
      space(),
      tipBox("「換気の話はロープレではなかなか出てこない知識」（先輩談）。詳しい説明は必要ないが、お客様から聞かれたときに「SUGOIEは熱を回収しながら換気する、RAKUIEはシンプルで信頼性の高い換気を採用しています」と答えられればOK。"),
      space(),
 
      // ===== 5. RAKUIE詳細 =====
      h1("5. RAKUIE（ラクイエ）の特徴と仕様"),
      space(),
      h2("5-1　RAKUIEの3つの柱"),
      space(),
      blueBox("RAKUIE（ラクイエ）のコンセプト", [
        "① タイムパフォーマンス：時短でも妥協のない家づくり。スピーディーに夢のマイホームを実現",
        "② コストパフォーマンス：最高の品質を最適な価格で。「手が届く贅沢」を感じられる住まい",
        "③ 安心安全の家づくり：地震に強い構造と省エネ設計で、長く安心して暮らせる住環境",
      ]),
      space(),
      h2("5-2　RAKUIEの性能等級"),
      space(),
      tbl(
        ["性能項目", "RAKUIE標準仕様の等級", "説明"],
        [
          ["断熱等性能等級", "等級5（ZEH水準）", "国の省エネ基準を大幅に上回る断熱性能"],
          ["一次エネルギー消費量等級", "等級6（最高等級）", "BEI = 0.8以下。太陽光は含まない"],
          ["耐震等級", "等級3（最高等級）", "建築基準法の1.5倍の力に対して倒壊しない"],
          ["劣化対策等級", "等級3", "75〜90年大規模改修不要レベルの耐久性"],
          ["維持管理対策等級", "等級3（最高等級）", "配管清掃・点検・補修がしやすい設計"],
        ],
        [2800, 2200, 4026]
      ),
      space(),
      h2("5-3　RAKUIEの建物仕様"),
      space(),
      tbl(
        ["項目", "RAKUIE標準仕様"],
        [
          ["断熱材（壁）", "現場発泡硬質ウレタンフォーム 100mm（ダルトフォーム）"],
          ["断熱材（屋根）", "現場発泡硬質ウレタンフォーム 210mm"],
          ["断熱材（床）", "サットフォーム 105mm（ビーズ法ポリスチレンフォーム）"],
          ["サッシ", "YKK AP エピソードⅡNEO（アルミ樹脂複合・Low-Eペアガラス・アルゴンガス入）"],
          ["基礎", "ベタ基礎工法（標準）＋キソパッキング工法"],
          ["外壁", "サイディング（KMEW：親水）"],
          ["屋根", "スレート屋根・ガルバリウム鋼板"],
          ["換気", "第3種24時間換気システム"],
          ["給湯", "エコキュート 370L（オール電化）"],
          ["制震", "なし（オプション）"],
          ["太陽光", "なし（オプション）"],
          ["保証", "設備プレミアム10年保証・建物長期保証20年"],
        ],
        [2800, 6226]
      ),
      space(),
      h2("5-4　長期優良住宅認定"),
      space(),
      body("RAKUIEは長期優良住宅の認定基準を標準仕様でクリアしています（別途申請費が必要）。認定されると様々な税制優遇が受けられます。"),
      space(),
      tbl(
        ["税制優遇", "内容"],
        [
          ["固定資産税の減税", "50%減税が5年間（一般住宅は3年間）"],
          ["住宅ローン控除", "借入限度額が最大5,000万円（子育て世帯・若者夫婦世帯）"],
          ["不動産取得税", "課税標準額から1,300万円控除（一般住宅は1,200万円）"],
          ["登録免許税", "保存登記0.1%（一般住宅0.15%）"],
        ],
        [2800, 6226]
      ),
      space(),
      h2("5-5　RAKUIEの価格"),
      space(),
      tbl(
        ["条件", "価格"],
        [
          ["29.30坪まで", "2,035万円（税抜）/ 2,238.5万円（税込）"],
          ["29.30坪超の場合", "1坪あたり46.5万円（税抜）/ 51.1万円（税込）追加"],
        ],
        [3000, 6026]
      ),
      space(),
      tipBox("RAKUIEは「セミオーダー・打ち合わせ効率化・ZEH水準の高性能」を両立した商品。「短期間でいい家を建てたい」「まず現実的な金額で家づくりをスタートしたい」お客様に最適な入口商品。"),
      space(),
 
      // ===== 6. SUGOIE詳細 =====
      h1("6. SUGOIE（すごい家）の特徴と仕様"),
      space(),
      h2("6-1　SUGOIEのコンセプト"),
      space(),
      body("SUGOIEは「高性能・高仕様・完全自由設計」を実現したSHOWA GROUPの主力ブランドです。性能もデザインもこだわり抜く方向けの商品です。"),
      space(),
      h2("6-2　SUGOIEの建物仕様"),
      space(),
      tbl(
        ["項目", "SUGOIE標準仕様"],
        [
          ["断熱材（壁）", "現場発泡硬質ウレタンフォーム 105mm"],
          ["断熱材（屋根）", "現場発泡硬質ウレタンフォーム 210mm"],
          ["サッシ", "YKK AP APW330（完全樹脂サッシ・Low-Eペアガラス・アルゴンガス入）"],
          ["断熱性能", "HEAT20 G2グレード（樹脂窓で大幅向上）"],
          ["外壁", "サイディング（KMEW：光セラ）※セルフクリーニング機能"],
          ["屋根", "スレート屋根・ガルバリウム鋼板"],
          ["換気", "ダクトレス熱交換型第1種24時間換気システム"],
          ["給湯", "エコキュート 370L（オール電化）"],
          ["太陽光発電", "3.64kW（標準装備）"],
          ["制震", "国土交通大臣認定耐力壁「Kダンパー」（地震性能20年保証）"],
          ["ZEH対応", "ZEHビルダー認定（ZEH28B-01230-CT）"],
        ],
        [2800, 6226]
      ),
      space(),
      h2("6-3　SUGOIEの価格"),
      space(),
      tbl(
        ["条件", "価格"],
        [
          ["29.30坪まで", "2,263万円（税抜）/ 2,489.3万円（税込）"],
          ["29.30坪超の場合", "1坪あたり54.1万円（税抜）/ 59.5万円（税込）追加"],
        ],
        [3000, 6026]
      ),
      space(),
      tipBox("SUGOIEはRAKUIEより約250万円（税抜）高い設定。主な差額の理由は「樹脂窓（APW330）・第1種熱交換換気・太陽光発電3.64kW・Kダンパー（制震）・光セラ外壁」の5点。この違いをお客様に説明できるようにする。"),
      space(),
 
      // ===== 7. SUGOIE vs RAKUIE 徹底比較 =====
      h1("7. SUGOIE vs RAKUIE 徹底比較"),
      space(),
      h2("7-1　仕様の主な違い"),
      space(),
      tbl(
        ["比較項目", "SUGOIE", "RAKUIE", "差のポイント"],
        [
          ["価格（29.30坪）", "約2,263万円（税抜）", "約2,035万円（税抜）", "約228万円の差"],
          ["窓（サッシ）", "樹脂サッシ APW330", "アルミ樹脂複合 エピソードⅡNEO", "断熱性能・結露のしにくさが違う"],
          ["換気システム", "第1種熱交換型（ダクトレス）", "第3種24時間換気", "熱ロス削減 vs シンプル&低コスト"],
          ["太陽光発電", "3.64kW 標準装備", "なし（オプション）", "光熱費・売電収入に影響"],
          ["制震装置", "Kダンパー 標準装備", "なし（オプション）", "地震被害軽減レベルが異なる"],
          ["外壁", "光セラ（セルフクリーニング）", "親水サイディング", "メンテナンスコストが異なる"],
          ["断熱等級", "HEAT20 G2グレード", "等級5（ZEH水準）", "SUGOIEがさらに上"],
          ["プラン自由度", "ほぼ完全自由設計", "セミオーダー", "RAKUIEは打ち合わせ回数に制限あり"],
          ["保証", "設備10年・建物20年", "設備10年・建物20年", "同じ"],
        ],
        [2200, 2200, 2000, 2626]
      ),
      space(),
      h2("7-2　どちらを提案するか：お客様タイプ別フロー"),
      space(),
      voiceBox("🎤 勉強会での先輩の言葉", "まずRAKUIEでエントリーして、ニーズに応じてアップグレードを提案する流れが多くなります。広告やCMに多く出てくるのはSUGOIEですが、実際の売れ筋・入口になりやすいのはRAKUIEです。RAKUIEでもモデルハウスはめちゃくちゃいいなっていう人は楽家で全然進めていくには問題ない。こっちが樹脂になってだいぶ違うよなってなってきたら、こっちの世界じゃんってなってくる。"),
      space(),
      tbl(
        ["お客様のタイプ", "推奨商品", "提案のポイント"],
        [
          ["まず予算感を確認したい", "RAKUIE", "月々の返済シミュレーションから入る"],
          ["光熱費を下げたい", "SUGOIE", "太陽光・第1種換気・高断熱の組み合わせで説明"],
          ["地震が不安", "SUGOIE", "Kダンパー（制震）の20年保証を強調"],
          ["メンテナンス費用を抑えたい", "SUGOIE", "光セラ外壁・樹脂窓で長期コストを説明"],
          ["早く家を建てたい", "RAKUIE", "タイムパフォーマンスのコンセプトを伝える"],
          ["デザインにこだわりたい", "SUGOIE / カタリエ", "完全自由設計の可能性をモデルハウスで体感"],
          ["コスト重視", "RAKUIE", "ZEH水準の高性能を「十分な性能」として伝える"],
        ],
        [2500, 1600, 4926]
      ),
      space(),
      warnBox("RAKUIEを提案する際に「性能が低い商品」という印象を与えてはいけない。RAKUIEも断熱等級5・耐震等級3・長期優良住宅対応の「非常に高性能な住宅」。「このエリアに十分すぎる性能で、コストパフォーマンスに優れた商品です」と伝える。"),
      space(),
 
      // ===== 8. 制震装置（Kダンパー） =====
      h1("8. 制震装置「Kダンパー」（SUGOIE）"),
      space(),
      body("SUGOIEには国土交通大臣認定の耐力壁「Kダンパー」が標準装備されています。耐震等級3はRAKUIEも取得していますが、Kダンパーは「耐震」に「制震」をプラスした設備です。"),
      space(),
      tbl(
        ["用語", "説明"],
        [
          ["耐震", "地震の力に構造で「耐える」。建物が壊れないようにする"],
          ["制震", "地震のエネルギーを「吸収・減衰」させる。揺れを小さくして被害を軽減"],
          ["免震", "建物と地盤を切り離して揺れを「伝えない」。コストが高い"],
        ],
        [2000, 7026]
      ),
      space(),
      voiceBox("🎤 勉強会での先輩の言葉", "耐震等級は1〜3まであってダンパー入れなかったら耐震等級3取れる。いや3は保つはずです。そやなこれは大人の関係ないはい。なくてもなれます。まあプラスアルファの精神ということ。"),
      space(),
      tipBox("Kダンパーは「地震性能20年保証」付き。耐震等級3だけでも法的には最高等級だが、「繰り返す地震にも強い」「被害が少なく修繕費が抑えられる」という観点でKダンパーのメリットを伝えると差別化になる。"),
      space(),
 
      // ===== 9. 確認テスト =====
      h1("9. 確認テスト・習得チェック"),
      space(),
      h2("9-1　理解度チェック"),
      space(),
      body("以下の項目を自分の言葉で説明できるか確認しましょう。できたらチェックを入れてください。"),
      space(),
      check("播磨エリアの気候特性と、それが性能提案にどう影響するかを30秒で説明できる"),
      space(),
      check("現場発泡断熱材の6つの特徴（省エネ・気密・遮音・ノンフロン・耐久・防カビ）を説明できる"),
      space(),
      check("第1種・第2種・第3種換気の違いを、お客様にわかりやすく説明できる"),
      space(),
      check("SUGOIEのダクトレス第1種換気が「カビのリスクが低い」理由を説明できる"),
      space(),
      check("第3種換気（RAKUIE）が「シンプルで信頼性が高い」理由を説明できる"),
      space(),
      check("SUGOIEとRAKUIEの価格差と、主な仕様の違い5点を即答できる"),
      space(),
      check("「予算を抑えたいお客様」と「性能重視のお客様」それぞれに向けた提案の入り口を説明できる"),
      space(),
      check("長期優良住宅の税制優遇（固定資産税・住宅ローン控除）の概要を説明できる"),
      space(),
      h2("9-2　ロープレ課題"),
      space(),
      body("次回の研修までに以下のシナリオを想定してロープレ練習しておきましょう。"),
      space(),
      tbl(
        ["No.", "シナリオ"],
        [
          ["①", "「SUGOIEとRAKUIEって何が違うんですか？値段の差はどこで出てるんですか？」"],
          ["②", "「換気システムって第1種と第3種があるって聞いたんですが、どっちがいいんですか？」"],
          ["③", "「このエリアって断熱性能どれくらい必要ですか？高ければ高いほどいいですか？」"],
          ["④", "「予算が限られてるんですが、RAKUIEって性能は大丈夫ですか？」"],
          ["⑤", "「太陽光ってほんとに元が取れるんですか？どれくらいで回収できますか？」"],
        ],
        [600, 8426]
      ),
      space(),
      tipBox("ロープレは「答えを暗記する」のではなく「お客様の不安や疑問に共感した上で説明する」練習。まずお客様の気持ちに寄り添い、次に「この辺の気候では…」「RAKUIEでも十分な性能で…」と前提を伝えてから説明する流れを身につけましょう。"),
      space(),
 
      // footer
      new Paragraph({
        children: [new TextRun({ text: "SHOWA GROUP　営業新入社員研修テキスト　商品・テクノロジー編　※内容は随時更新されます。最新のパンフレット・カタログと合わせてご確認ください。", font: "Arial", size: 17, color: "777777" })],
        spacing: { before: 200, after: 0 },
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" } }
      }),
    ]
  }]
});
 
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/mnt/user-data/outputs/商品テクノロジー研修テキスト.docx", buffer);
  console.log("Done!");
});
 
