/**
 * Front Cover generation - Most-front Cover ONLY (Inside Cover comes later).
 *
 * This page uses a DIFFERENT margin (4cm/4cm/3.2cm/3.2cm) than the rest of
 * the report. To support that, Front Cover gets its own Word section
 * (section index 0), and margin.js's applyUniformMargin is told to skip it
 * via a "protected section count" flag.
 *
 * Three lines, Arial Bold, centred:
 *   - UNIVERSITI POLY-TECH MALAYSIA  - size 22, black, first line (fixed text)
 *   - PROJECT TITLE                  - size 20, centre of the page
 *   - STUDENT FULL NAME              - size 20, last line
 *
 * The project title and student name are Word content controls (fill-in
 * fields). While the student hasn't filled them in, they show RED
 * placeholder text. As soon as the student clicks one and types, Word
 * replaces the placeholder with their text in BLACK (capitals, Arial Bold
 * 20) - set by the control's own run properties (w:sdtPr/w:rPr).
 *
 * Built once (idempotent via a marker bookmark).
 */

/* global Word, Office */

import { fillInLinesOoxml } from "./fill-in-fields";

const cmToPt = (cm) => cm * 28.346;

export const FRONT_COVER_MARGIN = {
  top: cmToPt(4),
  bottom: cmToPt(4),
  left: cmToPt(3.2),
  right: cmToPt(3.2),
};

// Header / footer distance from the page edge (Layout > Margins > Layout tab).
export const FRONT_COVER_HEADER_FOOTER = {
  headerFromTop: cmToPt(1.25),
  footerFromBottom: cmToPt(1),
};

// A4 = 21 cm x 29.7 cm. Word's default is often Letter (21.59 x 27.94 cm).
export const PAPER_SIZE = "A4";

// Sections 1 and 2 use the cover margin; everything after uses NORMAL_MARGIN.
const COVER_SECTION_COUNT = 2;

// "Normal" margin: 1 inch (2.54 cm) on every side.
export const NORMAL_MARGIN = {
  top: cmToPt(2.54),
  bottom: cmToPt(2.54),
  left: cmToPt(2.54),
  right: cmToPt(2.54),
};

const MARKER_BOOKMARK = "bmFrontCoverGenerated";
const PROTECTED_SECTIONS_KEY = "docToolkitFrontCoverSectionCount";
const BLACK = "#000000";

export function isFrontCoverGenerated() {
  return Office.context.document.settings.get(PROTECTED_SECTIONS_KEY) > 0;
}

// True only if the cover is really in the document right now (its marker
// bookmark is there). Unlike the saved setting above, this is false again
// after the cover is deleted or undone, or in a new document.
export async function frontCoverExists(context) {
  const marker = context.document.getBookmarkRangeOrNullObject(MARKER_BOOKMARK);
  marker.load("isNullObject");
  await context.sync();
  return !marker.isNullObject;
}

function markFrontCoverSections(count) {
  Office.context.document.settings.set(PROTECTED_SECTIONS_KEY, count);
  Office.context.document.settings.saveAsync();
}

// New paragraphs copy the previous paragraph's formatting (bold, size 20,
// red...), so the empty section paragraphs are reset to plain body text.
function resetToPlain(p) {
  p.font.name = "Arial";
  p.font.bold = false;
  p.font.size = 12;
  p.font.color = BLACK;
  p.alignment = "Left";
  p.spaceBefore = 0;
  p.spaceAfter = 0;
  return p;
}

// ---------------- Cover lines ----------------
// spaceBefore (points) places each line: the university on the first line,
// the title around the centre, the name near the bottom. Kept so a title
// that wraps onto 3 lines still fits on one page (4 cm top/bottom margin).
const COVER_LINES = [
  { text: "UNIVERSITI POLY-TECH MALAYSIA", size: 22, spaceBefore: 0 },
  { field: "Project Title", tag: "eformat-cover-title", placeholder: "PROJECT TITLE", size: 20, spaceBefore: 250, bookmark: "bmFcProjectTitle" },
  { field: "Student Full Name", tag: "eformat-cover-name", placeholder: "STUDENT FULL NAME", size: 20, spaceBefore: 200, bookmark: "bmFcFullName" },
];

// ---------------- keep the cover on ONE page ----------------
//
// The cover is one page in its own section; if a long title pushes the
// student name onto a second page, the sections (and the headers and page
// numbers after them) go wrong. So after every Set Text, the space above
// the title and above the name is shrunk just enough for everything to fit.
// Short titles keep the normal spacing (250 pt / 200 pt).

// Arial Bold letter widths (1/1000 of the font size).
const ARIAL_BOLD_WIDTH = {
  A: 722, B: 722, C: 722, D: 722, E: 667, F: 611, G: 778, H: 722, I: 278, J: 556, K: 722,
  L: 611, M: 833, N: 722, O: 778, P: 667, Q: 778, R: 722, S: 667, T: 611, U: 722, V: 667,
  W: 944, X: 667, Y: 667, Z: 611, " ": 278, "-": 333, ".": 278, ",": 278, "(": 333,
  ")": 333, "&": 722, "/": 278, ":": 333, "'": 238,
};
const textWidthPt = (text, size) =>
  [...text].reduce((w, c) => w + (ARIAL_BOLD_WIDTH[c] || (/\d/.test(c) ? 556 : 667)), 0) / 1000 * size;

// A4 width minus the 3.2 cm left and right cover margins.
const COVER_TEXT_WIDTH_PT = 595.3 - FRONT_COVER_MARGIN.left - FRONT_COVER_MARGIN.right;
// A4 height minus the 4 cm top and bottom margins, less a small safety gap.
const COVER_TEXT_HEIGHT_PT = 841.9 - FRONT_COVER_MARGIN.top - FRONT_COVER_MARGIN.bottom - 10;
const LINE_HEIGHT = (size) => size * 1.15; // single line spacing
const MIN_GAP_PT = 12;

/** How many lines `text` wraps to on the cover (Arial Bold, capitals). */
export function coverLineCount(text, size) {
  const words = String(text).toUpperCase().split(/\s+/).filter(Boolean);
  if (!words.length) return 1;
  let lines = 1;
  let current = "";
  words.forEach((word) => {
    const next = current ? `${current} ${word}` : word;
    if (current && textWidthPt(next, size) > COVER_TEXT_WIDTH_PT) {
      lines += 1;
      current = word;
    } else {
      current = next;
    }
  });
  return lines;
}

/**
 * The space above the title and above the name so the cover fits one page.
 * Returns { title, name, shrunk, fits }.
 */
export function coverSpacing(titleLines, nameLines) {
  const [uni, title, name] = COVER_LINES;
  const textHeight =
    LINE_HEIGHT(uni.size) + titleLines * LINE_HEIGHT(title.size) + nameLines * LINE_HEIGHT(name.size);
  const normal = title.spaceBefore + name.spaceBefore;
  const free = COVER_TEXT_HEIGHT_PT - textHeight;
  if (free >= normal) return { title: title.spaceBefore, name: name.spaceBefore, shrunk: false, fits: true };
  const scale = Math.max(free, 0) / normal;
  return {
    title: Math.max(MIN_GAP_PT, Math.floor(title.spaceBefore * scale)),
    name: Math.max(MIN_GAP_PT, Math.floor(name.spaceBefore * scale)),
    shrunk: true,
    fits: free >= 2 * MIN_GAP_PT,
  };
}

/** Re-spaces the cover after its text changed. Returns coverSpacing()'s result. */
export async function fitCoverOnOnePage(context) {
  const [, title, name] = COVER_LINES;
  const titleMark = context.document.getBookmarkRangeOrNullObject(title.bookmark);
  const nameMark = context.document.getBookmarkRangeOrNullObject(name.bookmark);
  titleMark.load("isNullObject");
  nameMark.load("isNullObject");
  await context.sync();
  if (titleMark.isNullObject || nameMark.isNullObject) return null; // older cover

  const titlePara = titleMark.paragraphs.getFirst();
  const namePara = nameMark.paragraphs.getFirst();
  titlePara.load("text");
  namePara.load("text");
  await context.sync();

  const spacing = coverSpacing(
    coverLineCount(titlePara.text.trim(), title.size),
    coverLineCount(namePara.text.trim(), name.size)
  );
  titlePara.spaceBefore = spacing.title;
  namePara.spaceBefore = spacing.name;
  await context.sync();
  return spacing;
}

export async function generateFrontCover(context) {
  const marker = context.document.getBookmarkRangeOrNullObject(MARKER_BOOKMARK);
  marker.load("isNullObject");
  await context.sync();
  if (!marker.isNullObject) {
    return { created: false };
  }

  // ---------------- Most-front Cover (one page, one section) ----------------
  // Built as Word XML so the title and name can be fill-in fields with a red
  // placeholder and black typed text. The university line is the first
  // paragraph of the document (no blank line above it), so the page never
  // spills onto a second page. A throwaway empty paragraph goes last: Word
  // merges the LAST inserted paragraph into the paragraph it lands on.
  context.document.body.insertOoxml(fillInLinesOoxml(COVER_LINES, 840100), Word.InsertLocation.start);
  await context.sync();

  const coverParas = context.document.body.paragraphs;
  coverParas.load("items");
  await context.sync();
  let cursor = coverParas.items[COVER_LINES.length - 1]; // the student name line

  // Bookmark the fill-in lines so the task pane's "Set Text" can find them.
  COVER_LINES.forEach((line, i) => {
    if (line.bookmark) coverParas.items[i].getRange("Content").insertBookmark(line.bookmark);
  });

  // Tag the first paragraph so this block is detectable next time.
  const firstPara = context.document.body.paragraphs.getFirst();
  firstPara.getRange().insertBookmark(MARKER_BOOKMARK);

  // ---------------- Split into 3 sections ----------------
  // Section 1: the cover above.
  // Section 2: a blank page with the SAME margin as the cover (for the
  //            Inside Cover later).
  // Section 3: everything after, with the normal 1 inch margin.
  // The empty paragraphs are created first, then a section break is put
  // after the cover and after section 2's paragraph.
  const section2Para = resetToPlain(cursor.insertParagraph("", "After"));
  resetToPlain(section2Para.insertParagraph("", "After")); // first line of section 3
  cursor.insertBreak(Word.BreakType.sectionNext, "After");
  section2Para.insertBreak(Word.BreakType.sectionNext, "After");
  await context.sync();

  const sections = context.document.sections;
  sections.load("items");
  await context.sync();

  // Paper size: EVERY section becomes A4. Set first, because changing paper
  // size after the margins can make Word re-adjust them. Sections added later
  // copy the section before them, so they stay A4 too.
  sections.items.forEach((section) => {
    section.pageSetup.paperSize = PAPER_SIZE;
  });
  await context.sync();

  sections.items.forEach((section, i) => {
    const setup = section.pageSetup;
    if (i < COVER_SECTION_COUNT) {
      // Sections 1 and 2: cover margin + cover header/footer distance.
      setup.headerDistance = FRONT_COVER_HEADER_FOOTER.headerFromTop;
      setup.footerDistance = FRONT_COVER_HEADER_FOOTER.footerFromBottom;
      setup.topMargin = FRONT_COVER_MARGIN.top;
      setup.bottomMargin = FRONT_COVER_MARGIN.bottom;
      setup.leftMargin = FRONT_COVER_MARGIN.left;
      setup.rightMargin = FRONT_COVER_MARGIN.right;
    } else {
      // Section 3 onward: back to the normal margin, 1 inch on every side.
      setup.topMargin = NORMAL_MARGIN.top;
      setup.bottomMargin = NORMAL_MARGIN.bottom;
      setup.leftMargin = NORMAL_MARGIN.left;
      setup.rightMargin = NORMAL_MARGIN.right;
    }
  });
  await context.sync();

  // Tell margin.js's applyUniformMargin to leave sections 1 and 2 alone.
  markFrontCoverSections(COVER_SECTION_COUNT);

  return { created: true, sections: sections.items.length };
}
