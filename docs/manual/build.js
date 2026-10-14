// Builds the user guide (.docx) plus an HTML twin printed to PDF, in Chinese (zh) or English (en).
//   python docs/manual/make_screenshots.py docs/manual/shots <sample-dir>   (synthetic data only)
//   node docs/manual/build.js zh docs/manual/shots docs/MarkItDown-Guide-zh-TW.docx guide-zh.html
//   node docs/manual/build.js en docs/manual/shots docs/MarkItDown-Guide-en.docx guide-en.html
//   chrome --headless --no-pdf-header-footer --print-to-pdf=docs/MarkItDown-Guide-<lang>.pdf guide-<lang>.html
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, ImageRun, LevelFormat, Footer, PageNumber, PageBreak,
} = require("docx");

const [LANG, SHOTS, OUT, OUT_HTML] = process.argv.slice(2);
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

function snippet(lines, note) {
  H.push(`<pre>${lines.map(esc).join("\n")}</pre>` + (note ? `<p class="note">${esc(note)}</p>` : ""));
  const mono = { ascii: "Consolas", hAnsi: "Consolas", eastAsia: "Microsoft JhengHei", cs: "Consolas" };
  const out = lines.map((t, i) => new Paragraph({
    children: [new TextRun({ text: t, font: mono, size: 17 })],
    shading: { type: ShadingType.CLEAR, fill: "F5F5F5", color: "auto" },
    spacing: { after: 0, before: i === 0 ? 80 : 0, line: 260 }, indent: { left: 160, right: 160 },
  }));
  out.push(new Paragraph({ spacing: { before: 60, after: 160 }, children: note ? [new TextRun({ text: note, font: FONT, size: 18, color: MUTED })] : [] }));
  return out;
}

function table(widths, rows) {
  const total = widths.reduce((a, c) => a + c, 0);
  H.push("<table><colgroup>" + widths.map((w) => `<col style="width:${(100 * w / total).toFixed(1)}%">`).join("") + "</colgroup>" + rows.map((cells, r) => "<tr>" + cells.map((c) => (r === 0 ? `<th>${html(c)}</th>` : `<td>${html(c)}</td>`)).join("") + "</tr>").join("") + "</table>");
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

const raw = (h, para) => (H.push(h), para);
// The text lives in content.<lang>.js; this file only lays it out.
const content = require(path.join(__dirname, `content.${LANG}.js`))({
  p, h1, h2, bullet, step, b, run, code, box, table, image, snippet, raw, esc, Paragraph, TextRun, PageBreak, FONT, ACCENT, MUTED, WARN_BG,
});
const children = content.children;

const doc = new Document({
  creator: content.creator, title: content.title, description: "v1.7.1",
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
    ...["steps", "rules", "project", "example"].map((reference) => ({ reference, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 360 } } } }] })),
  ] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
      run(content.footer[0], { size: 18, color: MUTED }),
      new TextRun({ children: [PageNumber.CURRENT], size: 18, color: MUTED, font: FONT }),
      run(content.footer[1], { size: 18, color: MUTED }),
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
  const footer = `"${content.footer[0].replace(/　/g, " ")}" counter(page) "${content.footer[1]}"`;
  const css = fs.readFileSync(path.join(__dirname, "guide.css"), "utf8").replace("__FOOTER__", footer);
  fs.writeFileSync(OUT_HTML, `<!doctype html><html lang="${content.lang}"><head><meta charset="utf-8"><title>${esc(content.title)}</title><style>${css}</style></head><body>${body}</body></html>`);
  console.log("wrote", OUT_HTML);
}
