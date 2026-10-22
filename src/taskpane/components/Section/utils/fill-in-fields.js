/**
 * Fill-in lines for the Front Cover and Title Page, built as Word XML.
 *
 * A "fill-in" line is a Word content control. Until the student fills it in
 * it shows a RED placeholder (e.g. PROJECT TITLE). When the student clicks it
 * and types, Word replaces the placeholder with their text in BLACK - taken
 * from the control's own run properties (w:sdtPr/w:rPr). A fixed line
 * (e.g. UNIVERSITI POLY-TECH MALAYSIA) is plain black text.
 *
 * line = { text }                             fixed black text
 *      | { field, tag, placeholder }          fill-in field
 *      + { size, spaceBefore }                Arial Bold size (pt), space above (pt)
 */

/* global Word */

const RED = "FF0000";
const BLACK = "000000";

const toTwips = (pt) => Math.round(pt * 20);

function escapeXml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// Arial Bold, capitals, in the given colour and size.
function runProps(color, size) {
  const hp = Math.round(size * 2);
  return (
    `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:b/><w:bCs/><w:caps/>` +
    `<w:color w:val="${color}"/><w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/></w:rPr>`
  );
}

function lineXml(line, id) {
  const black = runProps(BLACK, line.size);
  // pPr children must keep this order: spacing, jc, rPr.
  const pPr =
    `<w:pPr><w:spacing w:before="${toTwips(line.spaceBefore || 0)}" w:after="0" w:line="240" w:lineRule="auto"/>` +
    `<w:jc w:val="center"/>${black}</w:pPr>`;

  if (!line.field) {
    return `<w:p>${pPr}<w:r>${black}<w:t xml:space="preserve">${escapeXml(line.text)}</w:t></w:r></w:p>`;
  }

  return (
    `<w:p>${pPr}<w:sdt><w:sdtPr>${black}` +
    `<w:alias w:val="${escapeXml(line.field)}"/><w:tag w:val="${escapeXml(line.tag)}"/><w:id w:val="${id}"/>` +
    `<w:showingPlcHdr/></w:sdtPr>` +
    `<w:sdtContent><w:r>${runProps(RED, line.size)}<w:t xml:space="preserve">${escapeXml(line.placeholder)}</w:t></w:r></w:sdtContent>` +
    `</w:sdt></w:p>`
  );
}

/**
 * Word XML package for the given lines. A throwaway empty paragraph goes
 * last: Word merges the LAST inserted paragraph into the paragraph it lands
 * on, so this keeps every real line's own settings.
 * idBase keeps the content-control ids unique per page (cover vs title page).
 */
export function fillInLinesOoxml(lines, idBase) {
  const body = lines.map((line, i) => lineXml(line, idBase + i)).join("") + "<w:p/>";
  return (
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships></pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>` +
    `</pkg:xmlData></pkg:part></pkg:package>`
  );
}


/**
 * Fills one fill-in line from the task pane with `text` in black, Arial Bold
 * capitals at `size` pt. The line is found by its content-control TAG; if
 * that isn't there (pages made by an earlier version), by its BOOKMARK.
 * Returns false if the line isn't in the document (page not generated yet).
 */
export async function setFillInText(context, { bookmark, tag }, text, size) {
  let range = null;

  if (tag) {
    const controls = context.document.contentControls.getByTag(tag);
    controls.load("items");
    await context.sync();
    if (controls.items.length > 0) {
      range = controls.items[0].insertText(text, Word.InsertLocation.replace);
    }
  }

  if (!range && bookmark) {
    const marked = context.document.getBookmarkRangeOrNullObject(bookmark);
    marked.load("isNullObject");
    await context.sync();
    if (!marked.isNullObject) {
      range = marked.paragraphs.getFirst().getRange("Content").insertText(text, Word.InsertLocation.replace);
      range.insertBookmark(bookmark); // keep it findable for the next change
    }
  }

  if (!range) return false;

  range.font.name = "Arial";
  range.font.bold = true;
  range.font.allCaps = true;
  range.font.italic = false;
  range.font.size = size;
  range.font.color = "#000000";
  await context.sync();
  return true;
}

// ---------------- keeping a fill-in page on ONE page ----------------
//
// Long text (a 5-line project title, a 2-line name...) can push the last
// line onto the next page. After each Set Text, the space above each line
// is shrunk - all in proportion - just enough for everything to fit.
// Short text keeps the normal spacing.

// Arial Bold letter widths (1/1000 of the font size); fill-in lines are
// Arial Bold capitals.
const ARIAL_BOLD_WIDTH = {
  A: 722, B: 722, C: 722, D: 722, E: 667, F: 611, G: 778, H: 722, I: 278, J: 556, K: 722,
  L: 611, M: 833, N: 722, O: 778, P: 667, Q: 778, R: 722, S: 667, T: 611, U: 722, V: 667,
  W: 944, X: 667, Y: 667, Z: 611, " ": 278, "-": 333, ".": 278, ",": 278, "(": 333,
  ")": 333, "&": 722, "/": 278, ":": 333, "'": 238,
};
const textWidthPt = (text, size) =>
  [...text].reduce((w, c) => w + (ARIAL_BOLD_WIDTH[c] || (/\d/.test(c) ? 556 : 667)), 0) / 1000 * size;

/** How many lines `text` wraps to (Arial Bold capitals) in a column `widthPt` wide. */
export function wrappedLineCount(text, size, widthPt) {
  const words = String(text).toUpperCase().split(/\s+/).filter(Boolean);
  if (!words.length) return 1;
  let lines = 1;
  let current = "";
  words.forEach((word) => {
    const next = current ? `${current} ${word}` : word;
    if (current && textWidthPt(next, size) > widthPt) {
      lines += 1;
      current = word;
    } else {
      current = next;
    }
  });
  return lines;
}

/**
 * Space above each line so all lines fit in `heightPt`.
 *   lines: [{ size, spaceBefore }], lineCounts: lines each one wraps to,
 *   extraPt: other height on the page (e.g. a page-break line after).
 * Returns { gaps: [pt...], shrunk, fits }.
 */
export function fitGaps(lines, lineCounts, heightPt, { extraPt = 0, minGap = 6 } = {}) {
  const textHeight = lines.reduce((h, l, i) => h + lineCounts[i] * l.size * 1.15, 0) + extraPt;
  const normal = lines.reduce((g, l) => g + (l.spaceBefore || 0), 0);
  const free = heightPt - textHeight;
  if (free >= normal) return { gaps: lines.map((l) => l.spaceBefore || 0), shrunk: false, fits: true };
  const scale = Math.max(free, 0) / normal;
  return {
    gaps: lines.map((l) => Math.max(Math.min(minGap, l.spaceBefore || 0), Math.floor((l.spaceBefore || 0) * scale))),
    shrunk: true,
    fits: free >= lines.length * minGap,
  };
}

/**
 * Re-spaces the fill-in lines (each found by its bookmark) so the page
 * stays one page. Returns fitGaps()'s result, or null if a line is missing.
 */
export async function fitFillInPage(context, lines, { widthPt, heightPt, extraPt = 0 }) {
  const marks = lines.map((l) => {
    const r = context.document.getBookmarkRangeOrNullObject(l.bookmark);
    r.load("isNullObject");
    return r;
  });
  await context.sync();
  if (marks.some((m) => m.isNullObject)) return null;

  const paras = marks.map((m) => {
    const p = m.paragraphs.getFirst();
    p.load("text");
    return p;
  });
  await context.sync();

  const counts = lines.map((l, i) => wrappedLineCount(paras[i].text.trim(), l.size, widthPt));
  const result = fitGaps(lines, counts, heightPt, { extraPt });
  paras.forEach((p, i) => {
    p.spaceBefore = result.gaps[i];
  });
  await context.sync();
  return result;
}
