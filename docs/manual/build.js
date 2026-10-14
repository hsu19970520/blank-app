// Builds the user guide: docs/MarkItDown-Guide-zh-TW.docx, plus an HTML twin printed to PDF.
//   python docs/manual/make_screenshots.py docs/manual/shots <sample-dir>   (synthetic data only)
//   node docs/manual/build.js docs/manual/shots docs/MarkItDown-Guide-zh-TW.docx guide.html
//   chrome --headless --no-pdf-header-footer --print-to-pdf=docs/MarkItDown-Guide-zh-TW.pdf guide.html
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, ImageRun, LevelFormat, Footer, PageNumber, PageBreak,
} = require("docx");

const SHOTS = process.argv[2], OUT = process.argv[3], OUT_HTML = process.argv[4];
// Every helper also records HTML in document order, so the PDF says exactly what the .docx says.
const H = [];
const esc = (t) => String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const html = (children) => (typeof children === "string" ? esc(children) : children.map((c) => c._h || "").join(""));
const FONT = { ascii: "Segoe UI", hAnsi: "Segoe UI", eastAsia: "Microsoft JhengHei", cs: "Segoe UI" };
const ACCENT = "005FB8", MUTED = "5D5D5D", BORDER = "D0D0D0", HEAD_BG = "EEF3F9", NOTE_BG = "F4F7FB", WARN_BG = "FFF6E5";
const W = 9026; // A4 width minus 1" margins, in DXA

const run = (text, o = {}) => { const r = new TextRun({ text, font: FONT, ...o }); r._h = o.bold ? `<b>${esc(text)}</b>` : esc(text); return r; };
const p = (children, o = {}) => (H.push(`<p>${html(children)}</p>`), new Paragraph({ children: typeof children === "string" ? [run(children)] : children, spacing: { after: 120, line: 300 }, ...o }));
const h1 = (t) => (H.push(`<h1>${esc(t)}</h1>`), new Paragraph({ heading: HeadingLevel.HEADING_1, children: [run(t)], spacing: { before: 360, after: 160 } }));
const h2 = (t) => (H.push(`<h2>${esc(t)}</h2>`), new Paragraph({ heading: HeadingLevel.HEADING_2, children: [run(t)], spacing: { before: 240, after: 120 } }));
const bullet = (children, level = 0) => (H.push(`<li data-list="ul">${html(children)}</li>`), new Paragraph({ numbering: { reference: "bullets", level }, children: typeof children === "string" ? [run(children)] : children, spacing: { after: 80, line: 300 } }));
const step = (children, ref = "steps") => (H.push(`<li data-list="ol-${ref}">${html(children)}</li>`), new Paragraph({ numbering: { reference: ref, level: 0 }, children: typeof children === "string" ? [run(children)] : children, spacing: { after: 80, line: 300 } }));
const b = (t) => run(t, { bold: true });
const code = (t) => { const r = new TextRun({ text: t, font: { ascii: "Consolas", hAnsi: "Consolas", eastAsia: "Microsoft JhengHei" }, color: "1F1F1F", shading: { type: ShadingType.CLEAR, fill: "F0F0F0", color: "auto" } }); r._h = `<code>${esc(t)}</code>`; return r; };

function box(paras, fill = NOTE_BG, edge = ACCENT) {
  H.push(`<div class="box${fill === WARN_BG ? " warn" : ""}">` + paras.map((c) => `<p>${html(c)}</p>`).join("") + "</div>");
  return paras.map((children) => new Paragraph({
    children: typeof children === "string" ? [run(children)] : children,
    shading: { type: ShadingType.CLEAR, fill, color: "auto" },
    border: { left: { style: BorderStyle.SINGLE, size: 18, color: edge, space: 8 } },
    indent: { left: 160, right: 160 }, spacing: { after: 0, line: 300 },
  })).concat([new Paragraph({ children: [], spacing: { after: 120 } })]);
}

function table(widths, rows) {
  H.push("<table>" + rows.map((cells, r) => "<tr>" + cells.map((c) => (r === 0 ? `<th>${html(c)}</th>` : `<td>${html(c)}</td>`)).join("") + "</tr>").join("") + "</table>");
  const border = { style: BorderStyle.SINGLE, size: 4, color: BORDER };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new Table({
    width: { size: widths.reduce((a, c) => a + c, 0), type: WidthType.DXA }, columnWidths: widths,
    rows: rows.map((cells, r) => new TableRow({
      tableHeader: r === 0, cantSplit: true,
      children: cells.map((content, c) => new TableCell({
        width: { size: widths[c], type: WidthType.DXA }, borders,
        shading: r === 0 ? { type: ShadingType.CLEAR, fill: HEAD_BG, color: "auto" } : undefined,
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: typeof content === "string" ? [run(content, { bold: r === 0, size: 20 })] : content, spacing: { after: 0, line: 280 } })],
      })),
    })),
  });
}

function image(file, widthPx, caption) {
  const data = fs.readFileSync(path.join(SHOTS, file));
  const w = data.readUInt32BE(16), h = data.readUInt32BE(20);
  H.push(`<figure><img src="data:image/png;base64,${data.toString("base64")}" style="width:${Math.round(widthPx * 0.75)}pt"><figcaption>${esc(caption || "")}</figcaption></figure>`);
  const out = [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 },
    children: [new ImageRun({ type: "png", data, transformation: { width: widthPx, height: Math.round(widthPx * h / w) },
      altText: { title: caption, description: caption, name: file } })] })];
  if (caption) out.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [run(caption, { size: 18, color: MUTED })] }));
  return out;
}

const INSTRUCTIONS = [
  "我會上傳用「MarkItDown 轉換器」從 Excel 轉出的檔案。請照這個方式處理：",
  "1. 先只讀「00_索引與邏輯.md」。它列出每張工作表、每一欄的類型，以及程式精確算好的合計、最大值、最小值和公式邏輯；這些數字可以直接引用，不要自己加總大量數字。",
  "2. 能用索引回答就直接回答。需要明細時，告訴我要上傳哪一個資料檔（檔名和大約 token 數在索引最後的「資料檔」表），一次只要需要的那幾個，不要要求整包。",
  "3. 問題只跟某個料號、客戶，或只跟最新的預測版本有關時，提醒我用轉換器的關鍵字篩選或「只轉最新一版」重新轉一次，檔案會小很多。",
  "4. 資料是 Tab 分隔，第一欄是 Excel 列號（後面有 h 表示該列在 Excel 中隱藏），數值是完整精度。引用數字時請註明工作表和儲存格位置（例如 Summary!F12）。",
];

const raw = (h, para) => (H.push(h), para);
const children = [
  // ----- cover
  raw('<div class="cover"><div class="t1">MarkItDown 轉換器</div><div class="t2">使用說明書</div><div class="t3">把 Excel、Word、PDF 轉成 Claude 好讀、又省 token 的檔案</div><div class="meta">適用版本：v1.7.0<br>下載：https://github.com/hsu19970520/blank-app/releases/latest</div></div>',
    new Paragraph({ spacing: { before: 1800, after: 200 }, children: [run("MarkItDown 轉換器", { size: 56, bold: true, color: ACCENT })] })),
  new Paragraph({ spacing: { after: 120 }, children: [run("使用說明書", { size: 40, bold: true })] }),
  new Paragraph({ spacing: { after: 600 }, children: [run("把 Excel、Word、PDF 轉成 Claude 好讀、又省 token 的檔案", { size: 26, color: MUTED })] }),
  new Paragraph({ children: [run("適用版本：v1.7.0", { color: MUTED })], spacing: { after: 60 } }),
  new Paragraph({ children: [run("下載：https://github.com/hsu19970520/blank-app/releases/latest", { color: MUTED })], spacing: { after: 600 } }),
  ...box([
    [b("三句話看懂")],
    "1. 把 Excel 拖進「快速轉換面板」，轉好的檔案會變成一張卡片。",
    "2. 把卡片拖進 Claude 的對話框。大檔案的卡片是小小的「索引」，先只給這張。",
    "3. 省 token 靠的是「少上傳」：先給索引，需要明細時 Claude 會告訴你要哪個檔；預測瀑布表按「只轉最新一版」。",
  ]),
  raw('<div class="pb"></div>', new Paragraph({ children: [new PageBreak()] })),

  // ----- 1
  h1("1. 這個工具能做什麼"),
  p("MarkItDown 轉換器以 Microsoft 的開源工具 MarkItDown 為核心，把各種文件轉成 Markdown（.md）文字檔，方便交給 Claude Desktop 或 Cowork 閱讀。PDF、Word、PowerPoint、HTML、CSV、網頁和 YouTube 字幕等格式，轉換結果和原版相同。"),
  p([run("Excel 有特別處理，預設輸出「"), b("給 Claude 的精簡版"), run("」：")]),
  bullet([b("索引檔（00_索引與邏輯.md）"), run("：每張工作表的欄位，以及程式精確算好的合計、最大值、最小值、每欄的公式寫法、工作表之間的資料流向。通常只有幾千 token。")]),
  bullet([b("資料檔"), run("：完整、無損的資料。大檔案會依工作表拆成好幾個檔。")]),
  bullet([b("逐格核對"), run("：寫出檔案前，程式會把每一格讀回來和 Excel 比對，有一格不同就不產生結果。")]),
  ...box([
    [b("最重要的觀念")],
    "把大檔案「全部」上傳，花的 token 跟直接丟 Excel 差不多，因為數字本身就佔 token，換什麼格式都壓不小。真正省 token 的方法是「少上傳」：先給索引、只給需要的工作表、用關鍵字或「只轉最新一版」篩選。第 4 章有完整做法。",
  ], WARN_BG, "C27C0E"),

  // ----- 2
  h1("2. 下載與開啟"),
  step([run("到下載頁面下載 "), b("markitdown.exe"), run("（約 110 MB，不需要安裝，也不需要 Python）。")]),
  step([run("第一次執行時，Windows 可能會出現「Windows 已保護您的電腦」。這是因為程式沒有數位簽章，請按「"), b("其他資訊"), run("」，再按「"), b("仍要執行"), run("」。")]),
  step("雙擊就會打開主視窗。每次啟動需要幾秒鐘解壓縮，屬於正常現象。"),
  step([run("建議：在桌面建立 markitdown.exe 的捷徑，右鍵 › 內容，在「目標」最後面加上 "), code(" --panel"), run("。之後雙擊這個捷徑就會直接打開快速轉換面板。")]),

  // ----- 3
  h1("3. 快速轉換面板（建議的用法）"),
  p("快速轉換面板是一個浮在其他視窗上方的小視窗，可以把它放在 Claude 旁邊：檔案拖進來，結果拖出去。"),
  h2("3.1 開啟面板"),
  bullet([run("主視窗選單「檢視 › 快速轉換面板」，或按 "), code("Ctrl+Shift+P"), run("。")]),
  bullet("主視窗清單還是空的時候，按中間的藍色連結「或開啟快速轉換面板」。"),
  bullet([run("用加上 "), code("--panel"), run(" 的捷徑直接開啟（見第 2 章）。")]),
  h2("3.2 轉換檔案"),
  p("把檔案拖到面板上方的虛線框（可以一次拖好幾個），或按「選擇檔案…」。轉換中可以繼續拖檔案進來，會依序轉換。"),
  ...image("panel_empty.png", 260, "圖 1　面板剛打開時：把檔案拖到虛線框裡"),
  h2("3.3 把結果交給 Claude"),
  p([run("每個結果會變成一張卡片。"), b("把卡片直接拖進 Claude Desktop 的對話框"), run("，就會附加那個檔案。不方便拖曳時，在卡片上按兩下可以開啟檔案；按右鍵可以「複製內容」直接貼到 Claude，或「開啟檔案位置」從檔案總管拖。")]),
  ...image("panel_results.png", 280, "圖 2　轉換結果：上面是大檔案的索引卡片，下面是小檔案的卡片"),
  h2("3.4 看懂卡片"),
  table([2600, 6426], [
    ["卡片上的文字", "意思"],
    ["檔名", "拖進 Claude 的就是這個檔。大檔案是「00_索引與邏輯.md」。"],
    ["約 1,300 token", "這個檔大約要花多少 token（估計值，實際依模型可能差 1～2 成）。"],
    ["✓ 已逐格核對", "Excel 的每一格都已經和原檔比對過。"],
    ["先拖這張就好，Claude 會說還要哪個檔", "大檔案只要先給索引，Claude 需要明細時會告訴你檔名。"],
    ["同資料夾另有 N 個資料檔", "資料檔在 <檔名>_md 資料夾裡，需要時再上傳指定的那一個。"],
    ["⚠ 全部約 N 萬 token，太大，別整包上傳", "超過一次對話適合處理的量，不要把整個資料夾打包上傳。"],
    ["→ 只轉最新一版 · 約 N token", "預測瀑布表這類檔案才會出現（見 3.5）。按一下產生只有最新版本的小檔。"],
    ["來自 ○○.xlsx", "原始檔案。篩選出來的小檔會加註「只含最新 1 版」。"],
  ]),
  h2("3.5 只轉最新一版（預測瀑布表）"),
  p("有些 Excel 每週會累積一個新版本，例如預測瀑布表：同一個料號在每一版都有一列。程式會自動找出這種「版本欄」，在卡片上提供「只轉最新一版」、「只轉最新 2 版」的連結，並寫出大約多少 token。"),
  p("按一下連結，就會另外產生一個小檔（檔名後面加「_最新1版」），原本的完整版不會被覆蓋。只想問「這週的預測」或「這週和上週差多少」時，用這個小檔就夠了。"),
  ...image("panel_latest.png", 280, "圖 3　按「只轉最新一版」後，最上面多了一張只含最新版本的小檔卡片"),
  p([run("說明：PO 日期、出貨日期這類每天筆數不一樣的日期欄，不會被當成版本欄。只有大檔案（全部超過約 3 萬 token）才會出現這兩個連結。")], { }),
  h2("3.6 面板底部"),
  bullet([b("設定說明"), run("：目前的 Excel 輸出方式和儲存位置。")]),
  bullet([b("浮在其他視窗上方"), run("：預設開啟，方便把卡片拖到 Claude；不需要時可以關掉。")]),
  bullet([b("更多"), run("：「轉換選項…」會打開主視窗調整選項；「複製 Claude 專案指示」見第 5 章。")]),

  // ----- 4
  h1("4. 怎麼上傳給 Claude 最省 token"),
  table([3000, 3626, 2400], [
    ["你的狀況", "怎麼做", "大約上傳量"],
    ["小檔案（卡片上沒有「先拖這張就好」）", "直接拖卡片進 Claude", "就是卡片上的 token 數"],
    ["大檔案，問整體狀況、合計、有哪些欄位", "只拖索引卡片", "幾千 token"],
    ["大檔案，需要某張工作表的明細", "先給索引；Claude 說要哪個檔，再從 _md 資料夾上傳那一個", "每個資料檔最多約 10 萬"],
    ["預測瀑布表，問這週或這週和上週", "按「只轉最新一版」或「只轉最新 2 版」", "通常是全部的 1/10 以下"],
    ["只問某個料號或客戶", "在「轉換選項…」設定關鍵字篩選後重新轉", "通常幾千到幾萬"],
  ]),
  p("", { spacing: { after: 160 } }),
  p([b("五個原則")]),
  step([b("不要把整個資料夾打包上傳。"), run("這樣花的 token 跟直接丟 Excel 差不多。")], "rules"),
  step([b("不要叫 Claude「全部讀完再回答」。"), run("請它先看索引。")], "rules"),
  step([b("先給索引，需要時再補檔。"), run("索引最後的「資料檔」表列出每個檔大約多少 token。")], "rules"),
  step([b("能篩選就先篩選。"), run("只轉最新版本、只保留含關鍵字的列。")], "rules"),
  step([b("換主題就開新對話。"), run("同一個對話每問一次，前面的檔案都會重新算 token。")], "rules"),
  p([b("可以直接貼給 Claude 的開頭")], { spacing: { before: 200, after: 80 } }),
  ...box([
    "附件是 Excel 轉出的索引（00_索引與邏輯.md）。請先只根據索引回答；索引裡的合計和最大最小值是程式精確算好的，可以直接引用。如果需要明細，請告訴我要上傳哪一個資料檔（檔名和大約 token 數在索引最後），不要猜。",
    "我的問題是：……",
  ]),

  // ----- 5
  h1("5. 設定 Claude 專案指示（做一次就好）"),
  p("如果你常在同一個 Claude 專案（Project）裡處理 Excel，可以把固定指示貼到專案裡。之後在這個專案的每一次對話，Claude 都會先看索引、需要明細時才開口要檔案，也會提醒你用篩選。"),
  step([run("在面板按「"), b("更多 › 複製 Claude 專案指示"), run("」。面板底部會顯示「✓ 已複製」。")], "project"),
  step("在 Claude 建立一個專案（或打開現有的專案）。", "project"),
  step("打開專案的「指示」（Instructions）設定，貼上剛剛複製的內容，儲存。", "project"),
  step("之後在這個專案裡開對話，上傳索引卡片即可。", "project"),
  p([b("複製出來的指示內容如下：")], { spacing: { before: 160, after: 80 } }),
  ...box(INSTRUCTIONS),

  // ----- 6
  h1("6. 看懂轉出來的檔案"),
  h2("6.1 檔案放在哪裡"),
  bullet([b("小檔案"), run("：在原始檔旁邊產生一個同名的 .md。")]),
  bullet([b("大檔案"), run("：在原始檔旁邊產生資料夾「<檔名>_md」，裡面有 00_索引與邏輯.md 和 01_、02_… 開頭的資料檔；太大的工作表會再分成「_第1部分」、「_第2部分」。")]),
  bullet([b("篩選出來的小檔"), run("：檔名後面加「_最新1版」等，不會覆蓋完整版。")]),
  bullet("重新轉換同一個檔案時會更新原本的輸出；舊的分割檔會自動清掉，資料夾裡你自己放的其他檔案不會被動到。"),
  h2("6.2 索引裡有什麼"),
  bullet("每張工作表的狀態（顯示／隱藏）、列數、欄數、表頭在第幾列、公式數量、資料大約多少 token。"),
  bullet("每一欄的名稱、類型、非空格數，以及程式算好的最小值、最大值、合計。有「分類欄」（例如 Demand／ETD／Balance）時，合計會依分類分開算。"),
  bullet("每一欄的公式寫法、工作表之間的資料流向、自動篩選、隱藏的列和欄、資料驗證、樞紐分析表、版本欄。"),
  bullet("最後的「資料檔」表：每個資料檔的檔名和大約 token 數。"),
  h2("6.3 資料檔的格式"),
  bullet("Tab 分隔。第一行是 Excel 欄字母，之後每一行的第一欄是 Excel 列號；列號後面有 h 表示該列在 Excel 中是隱藏的（常見原因是被篩選掉）。"),
  bullet("數值是完整精度，沒有千分位；百分比用小數表示（0.25 = 25%）；日期寫成 YYYY-MM-DD，格式藏住的時間也會寫出來。"),
  bullet("儲存格的註解附在該工作表資料的最後面。"),

  // ----- 7
  h1("7. 主視窗與選項"),
  p("主視窗適合一次轉換很多檔案或整個資料夾，也可以預覽轉換結果。"),
  ...image("main_empty.png", 520, "圖 4　主視窗：上方加入檔案、資料夾或網址，按「轉換」"),
  p([run("主視窗右上角的「"), b("選項"), run("」選單：")]),
  table([3600, 5426], [
    ["選項", "說明"],
    ["儲存到：原始檔案所在的資料夾／其他資料夾", "轉換結果要存在哪裡。"],
    ["加入資料夾時包含子資料夾", "加入資料夾時，連子資料夾裡的檔案一起轉。"],
    ["略過已經有 .md 的檔案", "重跑一整個資料夾時，跳過已經轉過的檔案。"],
    ["Excel：給 Claude 的精簡版（建議）", "預設開啟。關掉會改成一個大的完整表格 .md，比較適合人看，但比較耗 token。"],
    ["Excel：只保留含關鍵字的列…", "輸入料號、客戶名稱等（多個用逗號隔開），只輸出含有這些文字的列。表頭會一起保留，索引裡的統計仍是整張表的結果。"],
    ["Excel：只輸出看得到的工作表、列和欄", "跳過隱藏或被篩選掉的資料，等於 Excel 畫面上看到的內容。"],
    ["Excel：在每個儲存格的值下方附上公式", "需要逐格查公式時使用，會自動改用完整表格版，檔案會大很多。"],
    ["Excel：轉換前先用 Excel 重新計算", "需要電腦上有 Microsoft Excel。在副本中重新計算公式、重新整理樞紐分析表後再轉換，原始檔不會被修改。"],
  ]),
  p("面板和主視窗使用同一組選項。", { spacing: { before: 120 } }),

  // ----- 8
  h1("8. 常見問題"),
  ...[
    ["Windows 說「已保護您的電腦」，不讓我開？", "按「其他資訊」再按「仍要執行」。程式沒有數位簽章，所以第一次會被提醒。"],
    ["卡片拖不進 Claude？", "在卡片上按右鍵「開啟檔案位置」，從檔案總管把檔案拖進 Claude；或按右鍵「複製內容」，直接貼到對話框。"],
    ["我把全部檔案都上傳了，token 沒有少很多？", "這是正常的。全部上傳的量跟直接丟 Excel 差不多，請照第 4 章「少上傳」的做法。"],
    ["數字準確嗎？", "每一格都和 Excel 逐格核對過；索引裡的合計和最大最小值也是程式算的。數值是 Excel 上次存檔時算好的結果；如果檔案在別的地方改過、還沒重新計算，可以開啟「轉換前先用 Excel 重新計算」。"],
    ["為什麼 token 數都寫「約」？", "這是估計值。不同的 Claude 模型計算方式略有不同，實際數字可能差 1～2 成，但用來比較大小已經足夠。"],
    ["掃描的 PDF 轉不出文字？", "掃描檔是圖片，沒有文字層，這個工具沒有文字辨識（OCR）功能。"],
    ["為什麼有些欄位沒有名稱？", "Excel 裡那一欄本身就沒有表頭（例如一個合併儲存格橫跨好幾欄）。程式不會自己猜名稱，可以從索引裡該欄的統計看出內容。"],
    ["檔案很大，轉換要多久？", "一般幾秒；幾十萬個公式的大檔大約幾十秒。轉換中可以繼續拖其他檔案進來排隊。"],
  ].flatMap(([q, a]) => [
    raw(`<p class="q">Q　${esc(q)}</p>`, new Paragraph({ spacing: { before: 160, after: 40 }, children: [run("Q　" + q, { bold: true, color: ACCENT })] })),
    p([run("A　" + a)]),
  ]),

  // ----- 9
  h1("9. 進階：命令列"),
  p("熟悉命令列的人可以直接在「命令提示字元」或 PowerShell 使用（完整選項請執行 markitdown.exe --help）："),
  table([4200, 4826], [
    ["指令", "說明"],
    [[code("markitdown.exe 報表.xlsx --out-dir 輸出")], "轉一個檔案到「輸出」資料夾"],
    [[code("markitdown.exe 報表.xlsx --xlsx-latest 1")], "只保留最新一版（預測瀑布表）"],
    [[code("markitdown.exe 報表.xlsx --xlsx-rows-containing PN-0001")], "只保留含這個料號的列（可重複指定）"],
    [[code("markitdown.exe 報表.xlsx --xlsx-format readable")], "改用完整表格版"],
    [[code("markitdown.exe D:\\資料夾 -r --out-dir D:\\輸出")], "整個資料夾（含子資料夾）"],
    [[code("markitdown.exe --panel")], "只開快速轉換面板"],
  ]),
];

const doc = new Document({
  creator: "MarkItDown 轉換器", title: "MarkItDown 轉換器 使用說明書", description: "v1.7.0",
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 32, bold: true, color: ACCENT }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0, keepNext: true, keepLines: true } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 26, bold: true, color: "1F1F1F" }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1, keepNext: true, keepLines: true } },
    ],
  },
  numbering: { config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 280 } } } }] },
    ...["steps", "rules", "project"].map((reference) => ({ reference, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 360 } } } }] })),
  ] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
      run("MarkItDown 轉換器 使用說明書 v1.7.0　·　第 ", { size: 18, color: MUTED }),
      new TextRun({ children: [PageNumber.CURRENT], size: 18, color: MUTED, font: FONT }),
      run(" 頁", { size: 18, color: MUTED }),
    ] })] }) },
    children,
  }],
});
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("wrote", OUT, buf.length); });
if (OUT_HTML) {
  // Group consecutive list items into <ul>/<ol> (each numbered list restarts like the .docx).
  let body = "", open = null;
  for (const h of H) {
    const m = h.match(/^<li data-list="([^"]+)">/);
    const kind = m ? m[1] : null;
    if (open && kind !== open) { body += open === "ul" ? "</ul>" : "</ol>"; open = null; }
    if (kind && !open) { body += kind === "ul" ? "<ul>" : "<ol>"; open = kind; }
    body += h;
  }
  if (open) body += open === "ul" ? "</ul>" : "</ol>";
  const css = fs.readFileSync(path.join(__dirname, "guide.css"), "utf8");
  fs.writeFileSync(OUT_HTML, `<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8"><title>MarkItDown 轉換器 使用說明書</title><style>${css}</style></head><body>${body}</body></html>`);
  console.log("wrote", OUT_HTML);
}
