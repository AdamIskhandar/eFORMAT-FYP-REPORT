/**
 * Table of Contents: generate, update, add subheading.
 *
 * CONTENTS layout (like the sample):
 *
 *                                                              Page
 *   DECLARATION                                                   4
 *   ...front-matter pages that exist...
 *
 *   CHAPTER 1      INTRODUCTION                                   9
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
 * addChapter              - adds the next chapter (CHAPTER 6 / TITLE) with a
 *                           "6.1 Introduction" section, before References.
 * addSubheading           - adds the next-numbered child heading under a
 *                           parent (under 3.3 -> 3.3.5; under 2.3.4 -> "b.").
 *                           Does NOT update the CONTENTS (Update does).
 *
 * NUMBERING RULE: numbers go to 3 digits at most (2.3.4). A heading under a
 * 3-digit heading is lettered instead - a., b., c. (never 2.3.4.1) - and
 * lettered headings are NOT listed in the CONTENTS. Only chapters and
 * headings up to 3 digits appear there.
 */

/* global Word */

const MARKER_BOOKMARK = "bmTocGenerated";
const TOC_HEAD_BOOKMARK = "bmTocHead"; // the "CONTENTS" heading
const TOC_END_BOOKMARK = "bmTocEnd"; // empty line that ends the rows
const HEADING_BOOKMARK_PREFIX = "bmReportHeading_";
const MAX_HEADING_LEVEL = 4;
const NEW_SUBHEADING_PLACEHOLDER = "Write your content here.";
// Deepest heading level listed in the CONTENTS (3 = 2.3.4). Level 4
// headings are lettered (a., b., ...) and left out of the CONTENTS.
const CONTENTS_MAX_LEVEL = 3;
const LETTER_LEVEL = 4;
// What follows the letter: "a." Change to ")" for "a)".
const LETTER_SUFFIX = ".";

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
// = 18pt. Each chapter starts on a new page, "CHAPTER 1" 50 mm from the top.
const LINE_PT = 18; // one 1.5-spaced line of 12pt text
const SUB_TITLE_INDENT_PT = 54; // 0.75": where "BACKGROUND" starts after "1.1"

const HEADING_FORMAT = {
  // CHAPTER 1 / (2 lines) / INTRODUCTION / (2 lines)
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

// Body content format (same as step 10 "Content"): Arial 11, justified,
// 1.5 spacing, black, from the margin.
const CONTENT_SIZE = 11;

function applyContentFormat(target) {
  target.font.name = "Arial";
  target.font.size = CONTENT_SIZE;
  target.font.bold = false;
  target.font.italic = false;
  target.font.allCaps = false;
  target.font.color = "#000000";
}

/**
 * The "Write your content here." line, as a plain paragraph (no box), set
 * in the content format: Arial 11, justified, 1.5 spacing, black. So when
 * the student selects it and types, their text is already in the right
 * format - and so is every paragraph they add with Enter.
 */
function formatPlaceholder(paragraph) {
  paragraph.styleBuiltIn = Word.BuiltInStyleName.normal;
  applyContentFormat(paragraph);
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
//   level 1 "1 INTRODUCTION"  -> "CHAPTER 1" <2 line breaks> "INTRODUCTION"
//   (chapters are numbered 1, 2, 3 - an older "CHAPTER III" becomes "CHAPTER 3")
//   level 2+ "1.1 Background" -> "1.1" <tab> "Background"
// Headings already in that form, and unnumbered ones (Appendix, References),
// are returned unchanged.
const CHAPTER_GAP = "\v\v\v"; // line break + 2 blank lines

// Gap ABOVE "CHAPTER 1", in blank 1.5 lines (18pt each), counted from the top
// margin. Word drops "space before" at the top of a new page, so the gap is
// made of blank lines inside the heading instead - Word always keeps those.
// 4 lines = 72pt, so CHAPTER 1 sits about 50 mm from the page top.
// Want more gap? Make this bigger. Less? Make it smaller.
const CHAPTER_TOP_LINES = 4;
const CHAPTER_TOP = "\v".repeat(CHAPTER_TOP_LINES);

export function reportHeadingText(text, level) {
  const t = text.trim();
  if (level === 1) {
    // "1.5 spacing x 2" between CHAPTER 1 and the title = 2 blank lines,
    // i.e. 3 line breaks. Headings already converted are re-spaced too.
    // CHAPTER_TOP is the gap above, as blank lines (see CHAPTER_TOP_LINES).
    const done = t.match(/^(CHAPTER\s+[IVXLCDM\d]+)\s*\v[\s\v]*(.+)$/is);
    if (done) return `${CHAPTER_TOP}${chapterNumberLabel(done[1])}${CHAPTER_GAP}${done[2].trim()}`;
    if (/^CHAPTER\b/i.test(t)) return `${CHAPTER_TOP}${chapterNumberLabel(t)}`;
    const m = t.match(/^(\d+)\.?\s+(.+)$/s);
    return m
      ? `${CHAPTER_TOP}CHAPTER ${parseInt(m[1], 10)}${CHAPTER_GAP}${m[2].trim()}`
      : `${CHAPTER_TOP}${t}`; // REFERENCES, APPENDIX ... get the same gap
  }
  const m = t.match(/^(\d+(?:\.\d+)+)\.?[ \t]+(.+)$/s);
  return m ? `${m[1]}\t${m[2].trim()}` : t;
}

// "CHAPTER III ..." -> "CHAPTER 3 ..." (chapters use numbers, not Roman).
function chapterNumberLabel(text) {
  return text.replace(/^CHAPTER\s+([IVXLCDM]+)\b/i, (_, roman) => `CHAPTER ${romanToInt(roman)}`);
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

// 1 -> "a", 2 -> "b", ... 26 -> "z", 27 -> "aa"
export function toLetter(n) {
  let out = "";
  let num = n;
  while (num > 0) {
    num -= 1;
    out = String.fromCharCode(97 + (num % 26)) + out;
    num = Math.floor(num / 26);
  }
  return out;
}

// Title of a level-4 heading without its old label: "2.3.4.1 Title",
// "a.<tab>Title", "b)<tab>Title" -> "Title". (A letter label only counts
// when a tab follows it, so a title like "Dr. Smith" is never cut.)
function letteredTitle(text) {
  return text
    .trim()
    .replace(/^\d+(?:\.\d+)+\.?\s+/, "")
    .replace(/^[a-z]{1,2}[.)]\t+/i, "")
    .trim();
}

const letterLabel = (n) => `${toLetter(n)}${LETTER_SUFFIX}`;

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

  // Under a 3-digit heading (2.3.4) the next one is a letter, not 2.3.4.1.
  if (parent.level + 1 === LETTER_LEVEL) {
    let count = 0;
    let nextIndex = -1;
    for (let j = parentIndex + 1; j < headings.length; j++) {
      if (headings[j].level <= parent.level) {
        nextIndex = j;
        break;
      }
      if (headings[j].level === LETTER_LEVEL) count++;
    }
    return { number: letterLabel(count + 1), level: LETTER_LEVEL, parentIndex, nextIndex };
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
  paragraphs.items.forEach((p, index) => {
    if (p.tableNestingLevel > 0) return;
    const level = headingLevelOf(p);
    const text = p.text.trim();
    if (!level || !text) return;
    headings.push({ paragraph: p, level, text, rawText: p.text, prefix: headingPrefix(text), index });
  });
  // Every paragraph of the document, for finding where a section ends.
  headings.all = paragraphs.items;
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

/**
 * Moves each front-matter bookmark onto the heading WORDS only (e.g. "LIST
 * OF FIGURES"). If the bookmark also covered a page break in front of the
 * heading, or another list's heading, its page number came out as the page
 * BEFORE (the List of Figures showed page 11 while it was on page 12).
 */
async function pinFrontMatterBookmarks(context, front) {
  // Search the WHOLE bookmark, not just its first paragraph: when a list is
  // inserted in front of another, Word stretches the first list's bookmark
  // over the new one (seen: the List of Figures bookmark began at LIST OF
  // TABLES, one page earlier).
  const found = front.map((f) => {
    const hits = context.document.getBookmarkRangeOrNullObject(f.bookmark).search(f.label, { matchCase: false });
    hits.load("items");
    return hits;
  });
  await context.sync();
  front.forEach((f, i) => {
    if (found[i].items.length > 0) found[i].items[0].insertBookmark(f.bookmark);
  });
  await context.sync();
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

const TITLE_COL_PT = inchToPt(1.4); // where titles start (after "CHAPTER 10")
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
    const label = `CHAPTER ${parseInt(e.number, 10)}`;
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

  // 0. Bring every heading into the report format (CHAPTER 1 / INTRODUCTION,
  //    "1.1<tab>BACKGROUND", Arial 12 bold, 1.5 spacing). Safe to repeat.
  //    Level-4 headings are (re)lettered a., b., c. under each 3-digit
  //    heading, in order - so a deleted or moved one never leaves a gap.
  const before = await scanHeadings(context);
  let letterCount = 0;
  before.forEach((h) => {
    let wanted;
    if (h.level === LETTER_LEVEL) {
      letterCount += 1;
      wanted = `${letterLabel(letterCount)}\t${letteredTitle(h.text)}`;
    } else {
      letterCount = 0; // a new 1/2/3-level heading restarts the letters
      wanted = reportHeadingText(h.text, h.level);
    }
    if (wanted !== h.rawText.replace(/[ \t\r\n]+$/, "")) {
      h.paragraph.getRange("Content").insertText(wanted, Word.InsertLocation.replace);
    }
    formatHeading(h.paragraph, h.level);
  });
  await context.sync();

  // 1. Bookmark every heading listed in the CONTENTS - chapters and up to
  //    3 digits (link + page number target). Lettered ones are left out.
  const headings = (await scanHeadings(context)).filter((h) => h.level <= CONTENTS_MAX_LEVEL);
  headings.forEach((h, i) => {
    h.paragraph.getRange("Content").insertBookmark(`${HEADING_BOOKMARK_PREFIX}${i}`);
  });
  await context.sync();

  // 2. The entries: front-matter pages that exist, then every heading.
  const front = await existingFrontMatter(context);
  await pinFrontMatterBookmarks(context, front);
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
    .filter((h) => h.prefix && h.level <= CONTENTS_MAX_LEVEL)
    // Chapter headings hold line breaks (CHAPTER 1 / INTRODUCTION) - show
    // them as one line: "CHAPTER 1 INTRODUCTION".
    .map((h) => ({ prefix: h.prefix, text: h.text.replace(/[\v\s]+/g, " ").trim(), level: h.level }));
}

/**
 * Adds "<parent>.<next> <title>" as the last child of the parent heading,
 * with a placeholder line under it. The Table of Contents is left for the
 * student to update (Update Table of Contents) when they're done adding.
 */
// A line that holds nothing but a page/section break or spaces.
function isBreakOrEmpty(paragraph) {
  return paragraph.tableNestingLevel === 0 && paragraph.text.replace(/[\f\s\u000c\u000e]/g, "") === "";
}

export async function addSubheading(context, parentPrefix, title) {
  const cleanTitle = (title || "").trim();
  if (!cleanTitle) throw new Error("Type a title for the new subheading.");

  const headings = await scanHeadings(context);
  const plan = planNewSubheading(headings, parentPrefix);
  const text = `${plan.number}\t${cleanTitle}`;

  // Put it straight after the parent's last line of content. NOT just
  // before the next heading: when the next heading is a new chapter (or
  // References / Appendix) there is a page break in between, and the new
  // subheading would land on the next chapter's page. So step back over
  // the page break and empty lines to the last real line.
  let newHeading;
  if (plan.nextIndex >= 0) {
    const all = headings.all;
    const parentIndex = headings[plan.parentIndex].index;
    let k = headings[plan.nextIndex].index - 1;
    // Step back over page breaks / empty lines. (A content box from an
    // earlier test version counts as content - stop there, go after it.)
    let box = null;
    for (; k > parentIndex; k--) {
      const parentBox = all[k].parentContentControlOrNullObject;
      parentBox.load("isNullObject");
      await context.sync();
      if (!parentBox.isNullObject) {
        box = parentBox;
        break;
      }
      if (!isBreakOrEmpty(all[k])) break;
    }
    if (!box && k > parentIndex) {
      const parentBox = all[k].parentContentControlOrNullObject;
      parentBox.load("isNullObject");
      await context.sync();
      if (!parentBox.isNullObject) box = parentBox;
    }
    // Inside a content box, go after the whole box (not inside it).
    newHeading = box ? box.insertParagraph(text, "After") : all[k].insertParagraph(text, "After");
  } else {
    newHeading = context.document.body.insertParagraph(text, "End");
  }
  formatHeading(newHeading, plan.level);

  const placeholder = newHeading.insertParagraph(NEW_SUBHEADING_PLACEHOLDER, "After");
  formatPlaceholder(placeholder);
  await context.sync();

  // The Table of Contents is NOT updated here - that takes a while, so the
  // student adds all their subheadings first, then clicks Update once.
  return { text: `${plan.number} ${cleanTitle}`, level: plan.level };
}

/**
 * Adds the next chapter after the last numbered chapter (so before
 * REFERENCES / APPENDIX), on its own page:
 *
 *   CHAPTER 6
 *   <TITLE>
 *   6.1   INTRODUCTION
 *   Write your content here.
 *
 * then updates the Table of Contents. Returns { number, title }.
 */
export async function addChapter(context, title) {
  const cleanTitle = (title || "").replace(/\s+/g, " ").trim();
  if (!cleanTitle) throw new Error("Type a title for the new chapter.");

  const headings = await scanHeadings(context);
  const chapters = headings.filter((h) => h.level === 1);
  if (chapters.length === 0) {
    throw new Error("Generate the Table of Contents first - it creates the chapters.");
  }

  let lastNumbered = -1;
  let highest = 0;
  chapters.forEach((h, i) => {
    const n = parseInt(h.prefix, 10);
    if (h.prefix && !Number.isNaN(n)) {
      lastNumbered = i;
      highest = Math.max(highest, n);
    }
  });
  const number = highest + 1;
  // First unnumbered chapter-level heading after the last chapter
  // (REFERENCES, APPENDIX ...). The new chapter goes just before it.
  const before = chapters.slice(lastNumbered + 1).find((h) => !h.prefix) || null;

  const headingText = reportHeadingText(`${number} ${cleanTitle.toUpperCase()}`, 1);
  const introText = reportHeadingText(`${number}.1 Introduction`, 2);

  let chapterPara;
  if (before) {
    // REFERENCES already starts on a new page, so the new chapter takes
    // that page and a page break after it pushes REFERENCES to the next.
    chapterPara = before.paragraph.insertParagraph(headingText, "Before");
  } else {
    context.document.body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
    chapterPara = context.document.body.insertParagraph(headingText, "End");
  }
  formatHeading(chapterPara, 1);

  const introPara = chapterPara.insertParagraph(introText, "After");
  formatHeading(introPara, 2);

  const placeholder = introPara.insertParagraph(NEW_SUBHEADING_PLACEHOLDER, "After");
  formatPlaceholder(placeholder);
  if (before) {
    placeholder.insertBreak(Word.BreakType.page, "After");
  }
  await context.sync();

  await updateTableOfContents(context);
  return { number, title: cleanTitle.toUpperCase() };
}

/**
 * Refreshes EVERY page number in the document that points at a heading or
 * caption (Table of Contents, List of Figures, List of Tables) - like
 * pressing F9 on all of them - and returns the numbers shown afterwards,
 * so the caller can repeat until they stop changing.
 */
export async function refreshAllPageNumbers(context) {
  const fields = context.document.body.fields;
  fields.load("items/code");
  await context.sync();
  const pageRefs = fields.items.filter((f) => /^\s*PAGEREF\b/i.test(f.code || ""));
  pageRefs.forEach((f) => f.updateResult());
  await context.sync();
  pageRefs.forEach((f) => f.result.load("text"));
  await context.sync();
  const pages = pageRefs.map((f) => f.result.text.trim());

  // Sanity check: every chapter / heading comes AFTER the Contents page, so
  // its page number must be bigger. If Word hasn't finished laying out the
  // pages yet, they all show the same early number - then this is false.
  const target = (f) => ((f.code || "").match(/PAGEREF\s+(\S+)/i) || [])[1] || "";
  const contentsIndex = pageRefs.findIndex((f) => target(f) === TOC_HEAD_BOOKMARK);
  //   Also: headings come in page order, and a report's headings can't all
  //   be on ONE page - if they are (e.g. every chapter shows 8), Word hadn't
  //   laid out the pages yet when the numbers were read.
  let ready = true;
  const headingPages = pageRefs
    .map((f, i) => (target(f).startsWith(HEADING_BOOKMARK_PREFIX) ? parseInt(pages[i], 10) : null))
    .filter((n) => n !== null);
  if (contentsIndex >= 0) {
    const contentsPage = parseInt(pages[contentsIndex], 10);
    if (headingPages.some((n) => !(n > contentsPage))) ready = false;
  }
  for (let i = 1; i < headingPages.length; i++) {
    if (headingPages[i] < headingPages[i - 1]) ready = false;
  }
  if (headingPages.length > 1 && headingPages[headingPages.length - 1] === headingPages[0]) {
    ready = false;
  }
  return { pages, ready };
}
