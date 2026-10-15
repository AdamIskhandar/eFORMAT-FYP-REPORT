/**
 * Table of Contents: generate, update, add subheading.
 *
 * CONTENTS layout (like the sample):
 *
 *                                                              Page
 *   DECLARATION                                                   4
 *   ...front-matter pages that exist...
 *
 *   CHAPTER I      INTRODUCTION                                   9
 *   1.1            Introduction                                   9
 *                  1.3.1  UPTM use a manual way ...              10
 *
 * One line per entry (tab stops, no table, no dotted leaders):
 *   - each line is a Word hyperlink to its heading's bookmark
 *     (Ctrl+Click jumps there)
 *   - each page number is a PAGEREF field to that same bookmark - the same
 *     thing Word's own Table of Contents uses inside - so WORD works out and
 *     keeps the number right (F9 or printing refreshes it too)
 *
 * generateTableOfContents - one-time setup: the CONTENTS page plus every
 *                           chapter/section heading from REPORT_STRUCTURE.
 * updateTableOfContents   - rebuilds the CONTENTS from the headings that
 *                           exist now (added / renamed / deleted headings).
 * listParentHeadings      - headings a new subheading can go under.
 * addSubheading           - adds the next-numbered child heading under a
 *                           parent (under 3.3 -> 3.3.5), then updates.
 */

/* global Word */

const MARKER_BOOKMARK = "bmTocGenerated";
const TOC_HEAD_BOOKMARK = "bmTocHead"; // the "CONTENTS" heading
const TOC_END_BOOKMARK = "bmTocEnd"; // empty line that ends the rows
const HEADING_BOOKMARK_PREFIX = "bmReportHeading_";
const MAX_HEADING_LEVEL = 4;
const NEW_SUBHEADING_PLACEHOLDER = "Write your content here.";

// Front-matter pages listed at the top of CONTENTS, in this order. Only the
// ones that exist in the document (bookmark found) are listed.
const FRONT_MATTER = [
  { label: "DECLARATION", bookmark: "bmDeclHeading" },
  { label: "ACKNOWLEDGEMENTS", bookmark: "bmAckHead" },
  { label: "ABSTRACT", bookmark: "bmAbstractHead" },
  { label: "CONTENTS", bookmark: TOC_HEAD_BOOKMARK },
  { label: "LIST OF TABLES", bookmark: "bmTableListHeading" },
  { label: "LIST OF FIGURES", bookmark: "bmFigureListHeading" },
];

// ---------- formatting of the real headings in the document ----------

// Body format (UPTM sample): Arial 12, 1.5 line spacing. One blank 1.5 line
// = 18pt. Each chapter starts on a new page, "CHAPTER I" 50 mm from the top.
const LINE_PT = 18; // one 1.5-spaced line of 12pt text
const SUB_TITLE_INDENT_PT = 54; // 0.75": where "BACKGROUND" starts after "1.1"

const HEADING_FORMAT = {
  // CHAPTER I / (2 lines) / INTRODUCTION / (2 lines)
  // (gap above comes from CHAPTER_TOP_LINES - Word ignores spaceBefore at a page top)
  1: { builtIn: "heading1", allCaps: true, alignment: "Centered", leftIndent: 0, firstLineIndent: 0, spaceBefore: 0, spaceAfter: LINE_PT },
  // 1.1   BACKGROUND
  2: { builtIn: "heading2", allCaps: true, alignment: "Left", leftIndent: SUB_TITLE_INDENT_PT, firstLineIndent: -SUB_TITLE_INDENT_PT, spaceBefore: LINE_PT, spaceAfter: LINE_PT },
  // 1.1.1 Globalization and Organizational Management
  3: { builtIn: "heading3", allCaps: false, alignment: "Left", leftIndent: SUB_TITLE_INDENT_PT, firstLineIndent: -SUB_TITLE_INDENT_PT, spaceBefore: LINE_PT, spaceAfter: LINE_PT },
  4: { builtIn: "heading4", allCaps: false, alignment: "Left", leftIndent: SUB_TITLE_INDENT_PT, firstLineIndent: -SUB_TITLE_INDENT_PT, spaceBefore: LINE_PT, spaceAfter: LINE_PT },
};

function formatHeading(paragraph, level) {
  const f = HEADING_FORMAT[level] || HEADING_FORMAT[1];
  // Style first (so the Table of Contents finds it), then the look on top
  // (Word's default heading styles are blue and use other fonts/sizes).
  paragraph.styleBuiltIn = Word.BuiltInStyleName[f.builtIn];
  paragraph.font.name = "Arial";
  paragraph.font.size = 12;
  paragraph.font.bold = true;
  paragraph.font.italic = false;
  paragraph.font.allCaps = f.allCaps;
  paragraph.font.color = "#000000";
  paragraph.alignment = f.alignment;
  paragraph.leftIndent = f.leftIndent;
  paragraph.firstLineIndent = f.firstLineIndent;
  paragraph.lineSpacing = LINE_PT;
  paragraph.spaceBefore = f.spaceBefore;
  paragraph.spaceAfter = f.spaceAfter;
}

function formatPlaceholder(paragraph) {
  // Body text: plain, justified, 1.5 spacing, from the margin. Grey italic
  // only because it's a hint the student replaces.
  paragraph.styleBuiltIn = Word.BuiltInStyleName.normal;
  paragraph.font.name = "Arial";
  paragraph.font.size = 12;
  paragraph.font.bold = false;
  paragraph.font.italic = true;
  paragraph.font.allCaps = false;
  paragraph.font.color = "#8a8a8a";
  paragraph.alignment = "Justified";
  paragraph.leftIndent = 0;
  paragraph.firstLineIndent = 0;
  paragraph.lineSpacing = LINE_PT;
  paragraph.spaceBefore = 0;
  paragraph.spaceAfter = 0;
}

function formatPlainLine(paragraph, { size = 12, bold = false, alignment = "Left" } = {}) {
  paragraph.styleBuiltIn = Word.BuiltInStyleName.normal;
  paragraph.font.name = "Arial";
  paragraph.font.size = size;
  paragraph.font.bold = bold;
  paragraph.font.italic = false;
  paragraph.font.allCaps = false;
  paragraph.font.color = "#000000";
  paragraph.alignment = alignment;
  paragraph.leftIndent = 0;
  paragraph.firstLineIndent = 0;
  paragraph.spaceBefore = 0;
  paragraph.spaceAfter = 0;
}

// ---------- pure helpers (no Word calls) ----------

function romanToInt(roman) {
  const v = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
  const r = roman.toUpperCase();
  let total = 0;
  for (let i = 0; i < r.length; i++) {
    const cur = v[r[i]];
    const next = v[r[i + 1]] || 0;
    total += cur < next ? -cur : cur;
  }
  return total;
}

// Chapter number of a heading: "CHAPTER I ..." / "CHAPTER 3 ..." / "3 METHODOLOGY"
// / "3.3 Phases ..." (number and title may be separated by a space or tab).
export function headingPrefix(text) {
  const t = text.trim();
  let m = t.match(/^CHAPTER\s+(\d+)\b/i);
  if (m) return m[1];
  m = t.match(/^CHAPTER\s+([IVXLCDM]+)\b/i);
  if (m) return String(romanToInt(m[1]));
  m = t.match(/^(\d+(?:\.\d+)*)\.?\s+\S/);
  return m ? m[1] : null;
}

// Text of a heading in the report format:
//   level 1 "1 INTRODUCTION"  -> "CHAPTER I" <2 line breaks> "INTRODUCTION"
//   level 2+ "1.1 Background" -> "1.1" <tab> "Background"
// Headings already in that form, and unnumbered ones (Appendix, References),
// are returned unchanged.
const CHAPTER_GAP = "\v\v\v"; // line break + 2 blank lines

// Gap ABOVE "CHAPTER I", in blank 1.5 lines (18pt each), counted from the top
// margin. Word drops "space before" at the top of a new page, so the gap is
// made of blank lines inside the heading instead - Word always keeps those.
// 4 lines = 72pt, so CHAPTER I sits about 50 mm from the page top.
// Want more gap? Make this bigger. Less? Make it smaller.
const CHAPTER_TOP_LINES = 4;
const CHAPTER_TOP = "\v".repeat(CHAPTER_TOP_LINES);

export function reportHeadingText(text, level) {
  const t = text.trim();
  if (level === 1) {
    // "1.5 spacing x 2" between CHAPTER I and the title = 2 blank lines,
    // i.e. 3 line breaks. Headings already converted are re-spaced too.
    // CHAPTER_TOP is the gap above, as blank lines (see CHAPTER_TOP_LINES).
    const done = t.match(/^(CHAPTER\s+[IVXLCDM\d]+)\s*\v[\s\v]*(.+)$/is);
    if (done) return `${CHAPTER_TOP}${done[1]}${CHAPTER_GAP}${done[2].trim()}`;
    if (/^CHAPTER\b/i.test(t)) return `${CHAPTER_TOP}${t}`;
    const m = t.match(/^(\d+)\.?\s+(.+)$/s);
    return m
      ? `${CHAPTER_TOP}CHAPTER ${toRoman(parseInt(m[1], 10))}${CHAPTER_GAP}${m[2].trim()}`
      : `${CHAPTER_TOP}${t}`; // REFERENCES, APPENDIX ... get the same gap
  }
  const m = t.match(/^(\d+(?:\.\d+)+)\.?[ \t]+(.+)$/s);
  return m ? `${m[1]}\t${m[2].trim()}` : t;
}

function headingTitle(text) {
  return text
    .trim()
    .replace(/^CHAPTER\s+(?:\d+|[IVXLCDM]+)\s*[:.-]?\s*/i, "")
    .replace(/^\d+(?:\.\d+)*\.?\s+/, "");
}

function toRoman(n) {
  const map = [[10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]];
  let out = "";
  let num = n;
  for (const [v, sym] of map) {
    while (num >= v) {
      out += sym;
      num -= v;
    }
  }
  return out || String(n);
}

const normText = (t) => String(t).replace(/\s+/g, " ").trim().toUpperCase();

/**
 * Where a new subheading goes, and its number.
 * @param {Array<{level:number, prefix:string|null}>} headings - in document order
 * @param {string} parentPrefix - e.g. "3.3"
 */
export function planNewSubheading(headings, parentPrefix) {
  const parentIndex = headings.findIndex((h) => h.prefix === parentPrefix);
  if (parentIndex === -1) throw new Error(`Heading ${parentPrefix} was not found in the document.`);

  const parent = headings[parentIndex];
  if (parent.level >= MAX_HEADING_LEVEL) {
    throw new Error(`"${parentPrefix}" is already the deepest heading level, so it can't have sub-headings.`);
  }

  const childPattern = new RegExp(`^${parentPrefix.replace(/\./g, "\\.")}\\.(\\d+)$`);
  let maxChild = 0;
  headings.forEach((h) => {
    const m = h.prefix && h.prefix.match(childPattern);
    if (m) maxChild = Math.max(maxChild, parseInt(m[1], 10));
  });

  // After everything inside the parent: before the next heading at the
  // parent's level or higher.
  let nextIndex = -1;
  for (let j = parentIndex + 1; j < headings.length; j++) {
    if (headings[j].level <= parent.level) {
      nextIndex = j;
      break;
    }
  }

  return { number: `${parentPrefix}.${maxChild + 1}`, level: parent.level + 1, parentIndex, nextIndex };
}

// ---------- document scanning ----------

function headingLevelOf(p) {
  const m = /^Heading\s*(\d)$/i.exec(p.styleBuiltIn || "") || /^Heading\s*(\d)/i.exec(p.style || "");
  if (!m) return 0;
  const level = parseInt(m[1], 10);
  return level >= 1 && level <= MAX_HEADING_LEVEL ? level : 0;
}

async function scanHeadings(context) {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items/text,items/style,items/styleBuiltIn,items/tableNestingLevel");
  await context.sync();

  const headings = [];
  paragraphs.items.forEach((p) => {
    if (p.tableNestingLevel > 0) return;
    const level = headingLevelOf(p);
    const text = p.text.trim();
    if (!level || !text) return;
    headings.push({ paragraph: p, level, text, rawText: p.text, prefix: headingPrefix(text) });
  });
  return headings;
}

async function existingFrontMatter(context) {
  const ranges = FRONT_MATTER.map((f) => {
    const r = context.document.getBookmarkRangeOrNullObject(f.bookmark);
    r.load("isNullObject");
    return r;
  });
  await context.sync();
  return FRONT_MATTER.filter((_, i) => !ranges[i].isNullObject);
}

// ---------- public API ----------

export async function generateTableOfContents(context, structure) {
  const marker = context.document.getBookmarkRangeOrNullObject(MARKER_BOOKMARK);
  marker.load("isNullObject");
  await context.sync();
  if (!marker.isNullObject) {
    return { created: false, count: structure.length };
  }

  // --- CONTENTS page: heading + an empty end line; rows go between them ---
  // The pages before it (Declaration, Acknowledgements, Abstract) end with
  // their own page break, leaving an empty last line on a fresh page. Only
  // add a page break if that's not the case - otherwise a blank page appears.
  const lastLine = context.document.body.paragraphs.getLast();
  lastLine.load("text");
  await context.sync();
  if (lastLine.text.trim() !== "") {
    context.document.body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
  }

  const heading = context.document.body.insertParagraph("CONTENTS", Word.InsertLocation.end);
  formatPlainLine(heading, { size: 12, bold: true, alignment: "Centered" });
  heading.font.allCaps = true;
  heading.spaceAfter = 18;
  heading.getRange("Content").insertBookmark(TOC_HEAD_BOOKMARK);

  const tocEnd = context.document.body.insertParagraph("", Word.InsertLocation.end);
  formatPlainLine(tocEnd, { size: 2 });
  tocEnd.getRange().insertBookmark(TOC_END_BOOKMARK);

  // --- The real headings (and placeholders) students write under ---
  tocEnd.insertBreak(Word.BreakType.page, Word.InsertLocation.after);
  await context.sync();

  structure.forEach((entry, i) => {
    // Every chapter (and Appendix / References) starts on a new page. The
    // first one already does (page break after the CONTENTS).
    if (entry.level === 1 && i > 0) {
      context.document.body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
    }
    const headingPara = context.document.body.insertParagraph(
      reportHeadingText(entry.text, entry.level),
      Word.InsertLocation.end
    );
    formatHeading(headingPara, entry.level);
    if (entry.placeholder) {
      const placeholderPara = context.document.body.insertParagraph(entry.placeholder, Word.InsertLocation.end);
      formatPlaceholder(placeholderPara);
    }
  });

  const trailing = context.document.body.insertParagraph("", Word.InsertLocation.end);
  formatPlainLine(trailing);
  await context.sync();

  // One bookmark around the whole generated block, only so a later click can
  // tell it has already been generated.
  heading.getRange("Start").expandTo(trailing.getRange("End")).insertBookmark(MARKER_BOOKMARK);
  await context.sync();

  // Insert Word's Table of Contents now that the headings exist.
  const { count } = await updateTableOfContents(context);
  return { created: true, count };
}

// ---------- CONTENTS ----------

const toTwips = (pt) => Math.round(pt * 20);
const inchToPt = (inch) => inch * 72;

const TITLE_COL_PT = inchToPt(1.4); // where titles start (after "CHAPTER III")
const LEVEL4_EXTRA_PT = inchToPt(0.4); // extra indent for 3.5.1.1-style lines
const PAGE_GAP_PT = inchToPt(0.5); // keep wrapped titles clear of the page numbers
// Right-aligned tab stop for the page numbers, at the right margin. Tab
// positions are measured from the left margin: A4 width minus the section 3
// margins (1.52" left, 1.2" right).
const PAGE_TAB_PT = 595.3 - inchToPt(1.52) - inchToPt(1.2);

// Font and spacing above each kind of line.
const LINE_FORMAT = {
  1: { size: 12, bold: true, spaceBefore: 18 }, // front matter + chapters
  2: { size: 12, bold: false, spaceBefore: 8 },
  3: { size: 11, bold: false, spaceBefore: 0 },
  4: { size: 11, bold: false, spaceBefore: 0 },
};


function escapeXml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function rPr({ bold = false, size = 12 } = {}) {
  const hp = size * 2;
  return (
    `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>` +
    (bold ? "<w:b/><w:bCs/>" : '<w:b w:val="0"/>') +
    `<w:color w:val="000000"/><w:u w:val="none"/><w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/></w:rPr>`
  );
}

const textRun = (text, fmt) => `<w:r>${rPr(fmt)}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
const tabRun = (fmt) => `<w:r>${rPr(fmt)}<w:tab/></w:r>`;

/**
 * One CONTENTS line. Each line type only gets the tab stops it uses, so the
 * last tab always lands on the page-number stop at the right margin.
 *   hanging: first line at the margin, wrapped lines under the title column
 *   level4:  one extra stop for the deeper 3.5.1.1-style lines
 *   anchor:  bookmark to link to (Ctrl+Click jumps there)
 */
function rowXml(runs, spaceBeforePt, { hanging = true, level4 = false, anchor = null } = {}) {
  const tabs =
    (level4 ? `<w:tab w:val="left" w:pos="${toTwips(TITLE_COL_PT + LEVEL4_EXTRA_PT)}"/>` : "") +
    `<w:tab w:val="right" w:pos="${toTwips(PAGE_TAB_PT)}"/>`;
  const ind = hanging
    ? `<w:ind w:left="${toTwips(TITLE_COL_PT)}" w:hanging="${toTwips(TITLE_COL_PT)}" w:right="${toTwips(PAGE_GAP_PT)}"/>`
    : `<w:ind w:left="0" w:firstLine="0" w:right="${toTwips(PAGE_GAP_PT)}"/>`;
  const body = anchor ? `<w:hyperlink w:anchor="${escapeXml(anchor)}" w:history="1">${runs}</w:hyperlink>` : runs;
  return (
    `<w:p><w:pPr><w:tabs>${tabs}</w:tabs>` +
    `<w:spacing w:before="${toTwips(spaceBeforePt)}" w:after="0" w:line="240" w:lineRule="auto"/>` +
    `${ind}<w:jc w:val="left"/></w:pPr>${body}</w:p>`
  );
}

function entryRowXml(e) {
  const page = (fmt) => tabRun(fmt) + textRun(e.page || "", fmt);
  const opts = { anchor: e.bookmark };

  if (e.kind === "front") {
    const f = LINE_FORMAT[1];
    return rowXml(textRun(e.label, f) + page(f), f.spaceBefore, { ...opts, hanging: false });
  }

  if (e.level === 1) {
    const f = LINE_FORMAT[1];
    if (!e.number) {
      // Appendix / References: no number column.
      return rowXml(textRun(e.title.toUpperCase(), f) + page(f), f.spaceBefore, { ...opts, hanging: false });
    }
    const label = `CHAPTER ${toRoman(parseInt(e.number, 10))}`;
    return rowXml(textRun(label, f) + tabRun(f) + textRun(e.title.toUpperCase(), f) + page(f), f.spaceBefore, opts);
  }
  if (e.level === 2) {
    const f = LINE_FORMAT[2];
    return rowXml(textRun(e.number || "", f) + tabRun(f) + textRun(e.title, f) + page(f), f.spaceBefore, opts);
  }
  // Levels 3 and 4: number + title together in the title column, smaller.
  const f = LINE_FORMAT[e.level] || LINE_FORMAT[4];
  const level4 = e.level >= 4;
  const lead = level4 ? tabRun(f) + tabRun(f) : tabRun(f);
  const label = e.number ? `${e.number}  ${e.title}` : e.title;
  return rowXml(lead + textRun(label, f) + page(f), f.spaceBefore, { ...opts, level4 });
}

function rowsOoxml(entries) {
  const bold = { bold: true, size: 12 };
  const header = rowXml(tabRun(bold) + textRun("Page", bold), 0, { hanging: false });
  const body = entries.map(entryRowXml).join("");
  // Throwaway last paragraph: Word merges the LAST inserted paragraph into
  // the line it's inserted before (losing its settings).
  return (
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships></pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${header}${body}<w:p/></w:body></w:document>` +
    `</pkg:xmlData></pkg:part></pkg:package>`
  );
}

// Exposed for testing outside Word.
export const __test = { rowsOoxml, headingTitle, toRoman };

// The end line of the CONTENTS = the LAST line inside the end bookmark. When
// lines are inserted just before a bookmark, Word can stretch the bookmark
// over them; using its last line (not its first) always finds the real end,
// so old lines are always cleared and the table is never duplicated.
function contentsEndLine(context) {
  return context.document.getBookmarkRangeOrNullObject(TOC_END_BOOKMARK).paragraphs.getLast();
}

function contentsBlock(context) {
  const head = context.document.getBookmarkRangeOrNullObject(TOC_HEAD_BOOKMARK);
  return head.paragraphs.getFirst().getRange("Whole").expandTo(contentsEndLine(context).getRange("Whole"));
}

// Deletes everything between CONTENTS and the end line: fields first, each
// as a whole (a field's lines can't be deleted one by one - Word refuses with
// GeneralException), except the hidden TC entries; then the lines.
async function clearContentsBlock(context) {
  const fields = contentsBlock(context).fields;
  fields.load("items/code");
  await context.sync();
  fields.items.forEach((f) => {
    if (!/^\s*TC\s/i.test(f.code || "")) f.delete();
  });
  await context.sync();

  const lines = contentsBlock(context).paragraphs;
  lines.load("items");
  await context.sync();
  for (let i = lines.items.length - 2; i >= 1; i--) {
    lines.items[i].delete();
  }
  await context.sync();
}

/**
 * Rebuilds the CONTENTS from the document as it is now (see top of file).
 * Returns { count } = number of entries.
 */
export async function updateTableOfContents(context) {
  const head = context.document.getBookmarkRangeOrNullObject(TOC_HEAD_BOOKMARK);
  const end = context.document.getBookmarkRangeOrNullObject(TOC_END_BOOKMARK);
  head.load("isNullObject");
  end.load("isNullObject");
  await context.sync();
  if (head.isNullObject || end.isNullObject) {
    throw new Error(
      "Generate the Table of Contents first. (A Contents page made by an older version can't be updated - delete it and generate again.)"
    );
  }

  // 0. Bring every heading into the report format (CHAPTER I / INTRODUCTION,
  //    "1.1<tab>BACKGROUND", Arial 12 bold, 1.5 spacing). Safe to repeat.
  const before = await scanHeadings(context);
  before.forEach((h) => {
    const wanted = reportHeadingText(h.text, h.level);
    if (wanted !== h.rawText.replace(/[ \t\r\n]+$/, "")) {
      h.paragraph.getRange("Content").insertText(wanted, Word.InsertLocation.replace);
    }
    formatHeading(h.paragraph, h.level);
  });
  await context.sync();

  // 1. Bookmark every heading in document order (link + page number target).
  const headings = await scanHeadings(context);
  headings.forEach((h, i) => {
    h.paragraph.getRange("Content").insertBookmark(`${HEADING_BOOKMARK_PREFIX}${i}`);
  });
  await context.sync();

  // 2. The entries: front-matter pages that exist, then every heading.
  const front = await existingFrontMatter(context);
  const entries = [
    ...front.map((f) => ({ kind: "front", label: f.label, bookmark: f.bookmark })),
    ...headings.map((h, i) => ({
      kind: "heading",
      level: h.level,
      number: h.prefix,
      title: headingTitle(h.text),
      bookmark: `${HEADING_BOOKMARK_PREFIX}${i}`,
    })),
  ];

  // 3. The layout (lines + links), page column empty for now.
  await insertLayout(context, entries);

  // 4. Page numbers: a PAGEREF field at the end of each line. Added last, so
  //    every line is in place first; then refreshed once more so Word uses
  //    the final page layout.
  const lines = contentsBlock(context).paragraphs;
  lines.load("items");
  await context.sync();
  const FIRST = 2; // lines: [CONTENTS, "Page", ...entries, end line]
  const fields = [];
  entries.forEach((e, i) => {
    const line = lines.items[FIRST + i];
    if (!line) return;
    fields.push(line.getRange("End").insertField(Word.InsertLocation.end, Word.FieldType.pageRef, e.bookmark, false));
  });
  await context.sync();

  try {
    fields.forEach((f) => f.updateResult());
    await context.sync();
  } catch (err) {
    console.warn("Could not refresh the page numbers a second time (press F9 in Word if one looks off).", err);
  }

  return { count: entries.length };
}

/**
 * Refreshes the page-number fields in the CONTENTS (like pressing F9 on them).
 * Call this in a SEPARATE Word.run a moment after Update: Word finishes laying
 * out the pages in the background after the add-in's changes, and fields
 * refreshed before that all get the same, too-early page number.
 * Returns how many page numbers were refreshed.
 */
export async function refreshContentsPageNumbers(context) {
  const end = context.document.getBookmarkRangeOrNullObject(TOC_END_BOOKMARK);
  end.load("isNullObject");
  await context.sync();
  if (end.isNullObject) return 0;

  const fields = contentsBlock(context).fields;
  fields.load("items/code");
  await context.sync();
  const pageRefs = fields.items.filter((f) => /^\s*PAGEREF\b/i.test(f.code || ""));
  pageRefs.forEach((f) => f.updateResult());
  await context.sync();
  return pageRefs.length;
}

async function insertLayout(context, entries) {
  await clearContentsBlock(context);
  contentsEndLine(context).getRange("Start").insertOoxml(rowsOoxml(entries), "Before");
  await context.sync();

  // Pin the end bookmark back onto the end line only.
  contentsEndLine(context).getRange("Whole").insertBookmark(TOC_END_BOOKMARK);
  await context.sync();
}

/**
 * Numbered headings that can take a new subheading (deepest level excluded).
 */
export async function listParentHeadings(context) {
  const headings = await scanHeadings(context);
  return headings
    .filter((h) => h.prefix && h.level < MAX_HEADING_LEVEL)
    .map((h) => ({ prefix: h.prefix, text: h.text, level: h.level }));
}

/**
 * Adds "<parent>.<next> <title>" as the last child of the parent heading,
 * with a placeholder line under it, then updates the Table of Contents.
 */
export async function addSubheading(context, parentPrefix, title) {
  const cleanTitle = (title || "").trim();
  if (!cleanTitle) throw new Error("Type a title for the new subheading.");

  const headings = await scanHeadings(context);
  const plan = planNewSubheading(headings, parentPrefix);
  const text = `${plan.number}\t${cleanTitle}`;

  let newHeading;
  if (plan.nextIndex >= 0) {
    newHeading = headings[plan.nextIndex].paragraph.insertParagraph(text, "Before");
  } else {
    newHeading = context.document.body.insertParagraph(text, "End");
  }
  formatHeading(newHeading, plan.level);

  const placeholder = newHeading.insertParagraph(NEW_SUBHEADING_PLACEHOLDER, "After");
  formatPlaceholder(placeholder);
  await context.sync();

  await updateTableOfContents(context);
  return { text: `${plan.number} ${cleanTitle}`, level: plan.level };
}
