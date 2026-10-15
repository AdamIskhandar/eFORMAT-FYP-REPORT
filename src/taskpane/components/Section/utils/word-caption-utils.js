/**
 * Caption Automation Module (Task 4.1)
 * - Figure / table caption generation
 * - Chapter-based auto numbering: the first figure in Chapter 3 is
 *   "Figure 3.1", the next "Figure 3.2"; tables count separately
 *   ("Table 3.1", ...). The count restarts in every chapter.
 * - Caption validation rules
 *
 * Captions are plain paragraphs, "Figure 3.1: text" / "Table 3.1: text",
 * each anchored with a bookmark (Figure_3_1 / Table_3_1 - bookmark names
 * can't contain dots). Cross-references (word-crossref-utils.js) and the
 * List of Figures / Tables (caption-list-utils.js) find captions through
 * these bookmarks.
 *
 * The chapter of a caption is the last chapter heading above it: a
 * "Heading 1" paragraph or a line starting "CHAPTER 3" / "CHAPTER III".
 * A Heading 1 with no chapter number (REFERENCES, APPENDIX ...) ends the
 * chapters, so captions after it are numbered 1, 2, 3 ...
 *
 * Inserting a caption renumbers every caption of that type, so adding a
 * figure in the middle of Chapter 3 shifts the later ones (3.2 -> 3.3) and
 * updates their cross-references.
 */

/* global Word */

const CAPTION_CONFIG = {
  figure: { label: "Figure", bookmarkPrefix: "Figure_" },
  table: { label: "Table", bookmarkPrefix: "Table_" },
};

const NEW_CAPTION_BOOKMARK = "bmNewCaptionTemp";

function getConfig(type) {
  const config = CAPTION_CONFIG[type];
  if (!config) throw new Error(`Unknown caption type "${type}". Use "figure" or "table".`);
  return config;
}

// "3.1" -> "Figure_3_1"
export function numberToBookmark(type, number) {
  return `${getConfig(type).bookmarkPrefix}${String(number).replace(/\./g, "_")}`;
}

// "Figure_3_1" -> "3.1" (null if it isn't a caption bookmark of this type)
export function bookmarkToNumber(type, bookmarkName) {
  const { bookmarkPrefix } = getConfig(type);
  if (!bookmarkName || !bookmarkName.startsWith(bookmarkPrefix)) return null;
  const rest = bookmarkName.slice(bookmarkPrefix.length);
  return /^\d+(?:_\d+)*$/.test(rest) ? rest.replace(/_/g, ".") : null;
}

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

const isHeading1 = (p) =>
  /^Heading1$/i.test(p.styleBuiltIn || "") || /^Heading\s*1$/i.test(p.style || "");

// Returns undefined if the paragraph isn't a chapter-level heading, null for
// an unnumbered one (REFERENCES, APPENDIX), or the chapter number.
function chapterOfHeading(p) {
  if (p.tableNestingLevel > 0) return undefined;
  const t = p.text.trim(); // trim() also drops the line breaks above "CHAPTER I"
  let m = t.match(/^CHAPTER\s+(\d+)\b/i);
  if (m) return parseInt(m[1], 10);
  m = t.match(/^CHAPTER\s+([IVXLCDM]+)\b/i);
  if (m) return romanToInt(m[1]);
  if (!isHeading1(p)) return undefined;
  m = t.match(/^(\d+)\.?\s+\S/); // "3 METHODOLOGY"
  return m ? parseInt(m[1], 10) : null;
}

/**
 * All captions of a type, in document order - in normal paragraphs and in
 * table cells. A caption is a paragraph starting "Figure <number>" followed
 * by ":" "." or "-" (or any centered "Figure <number> ..." line). Each comes
 * back with the number it SHOULD have (`expected`) for its chapter.
 */
export async function getExistingCaptions(context, type) {
  const { label } = getConfig(type);
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load(
    "items/text,items/style,items/styleBuiltIn,items/alignment,items/tableNestingLevel"
  );
  await context.sync();

  const pattern = new RegExp(`^${label}\\s+(\\d+(?:\\.\\d+)*)\\s*([:.\\-\\u2013]?)\\s*(.*)$`, "is");

  const captions = [];
  let chapter = null;
  let region = 0; // changes at every chapter-level heading
  const counters = {};

  paragraphs.items.forEach((p, index) => {
    const ch = chapterOfHeading(p);
    if (ch !== undefined) {
      chapter = ch;
      region++;
      return;
    }

    const text = p.text.trim();
    const m = text.match(pattern);
    if (!m) return;
    const hasSeparator = m[2] !== "";
    if (!hasSeparator && p.alignment !== "Centered") return; // "Figure 3.1 shows ..." is a sentence
    if (/\t\s*\d+\s*$/.test(p.text)) return; // an entry of an old-style list (text <tab> page)

    counters[region] = (counters[region] || 0) + 1;
    const expected = chapter ? `${chapter}.${counters[region]}` : `${counters[region]}`;
    captions.push({
      paragraph: p,
      index,
      chapter,
      number: m[1],
      expected,
      separator: m[2],
      title: m[3].trim(),
      text,
    });
  });

  return captions;
}

/**
 * Renumbers every caption of a type by chapter (Figure 3.1, 3.2 ...) and
 * moves the bookmarks to match. Returns { oldNumber: newNumber } for every
 * caption, for updateReferencesAfterRenumber in word-crossref-utils.js.
 */
export async function renumberCaptions(context, type) {
  const { label, bookmarkPrefix } = getConfig(type);
  const captions = await getExistingCaptions(context, type);

  const numberMap = {};
  captions.forEach((c) => {
    if (!(c.number in numberMap)) numberMap[c.number] = c.expected;
  });

  // Replace only the "Figure 2" part, so the caption's own text and
  // formatting stay as they are.
  const searches = captions.map((c) => {
    if (c.number === c.expected) return null;
    const res = c.paragraph.search(`${label} ${c.number}`, { matchCase: false });
    res.load("items");
    return res;
  });
  await context.sync();
  searches.forEach((res, i) => {
    if (res && res.items.length > 0) {
      res.items[0].insertText(`${label} ${captions[i].expected}`, Word.InsertLocation.replace);
    }
  });
  await context.sync();

  // Bookmark each caption under its new number. Re-using a name moves that
  // bookmark, so going through in order leaves every name on the right
  // caption; leftovers (deleted captions, old numbers) are removed after.
  const wanted = new Set();
  captions.forEach((c) => {
    const name = numberToBookmark(type, c.expected);
    wanted.add(name);
    c.paragraph.getRange("Content").insertBookmark(name);
  });
  await context.sync();

  const all = context.document.body.getRange("Whole").getBookmarks(true, true);
  await context.sync();
  all.value
    .filter((name) => name.startsWith(bookmarkPrefix) && !wanted.has(name))
    .forEach((name) => context.document.deleteBookmark(name));
  await context.sync();

  return numberMap;
}

/**
 * Inserts a new caption at the cursor and numbers it for its chapter.
 * - Cursor in a table: a Table caption goes above the table, a Figure
 *   caption below it.
 * - Cursor anywhere else (a paragraph of text, a picture): the caption goes
 *   on a new line after that paragraph.
 * Every caption of the type is then renumbered, so later ones shift.
 */
export async function insertCaption(context, type, captionText) {
  const { label } = getConfig(type);
  const text = `${label} 0: ${captionText.trim()}`; // real number set below

  const selection = context.document.getSelection();
  const table = selection.parentTableOrNullObject;
  table.load("isNullObject");
  await context.sync();

  let inserted;
  if (!table.isNullObject) {
    inserted = table.insertParagraph(text, type === "table" ? "Before" : "After");
  } else {
    inserted = selection.paragraphs.getLast().insertParagraph(text, Word.InsertLocation.after);
  }

  // A plain caption line, whatever the paragraph it came from looked like
  // (a heading, a bullet, an indented paragraph).
  inserted.styleBuiltIn = Word.BuiltInStyleName.normal;
  inserted.load("isListItem");
  await context.sync();
  if (inserted.isListItem) inserted.detachFromList();
  inserted.font.name = "Arial";
  inserted.font.size = 10;
  inserted.font.bold = true;
  inserted.font.italic = true;
  inserted.font.color = "#000000";
  inserted.alignment = Word.Alignment.centered;
  inserted.leftIndent = 0;
  inserted.firstLineIndent = 0;
  inserted.spaceBefore = 6;
  inserted.spaceAfter = 12;
  inserted.getRange("Content").insertBookmark(NEW_CAPTION_BOOKMARK);
  await context.sync();

  const numberMap = await renumberCaptions(context, type);

  const mine = context.document.getBookmarkRangeOrNullObject(NEW_CAPTION_BOOKMARK);
  mine.load("isNullObject,text");
  await context.sync();
  const finalText = mine.isNullObject ? text : mine.text.trim();
  context.document.deleteBookmark(NEW_CAPTION_BOOKMARK);
  await context.sync();

  const m = finalText.match(new RegExp(`^${label}\\s+(\\d+(?:\\.\\d+)*)`, "i"));
  const number = m ? m[1] : null;
  return {
    number,
    bookmarkName: number ? numberToBookmark(type, number) : null,
    text: finalText,
    numberMap,
  };
}

/**
 * Checks every caption has the number its chapter gives it, no duplicates,
 * and text after the number.
 */
export async function validateCaptions(context, type) {
  const { label } = getConfig(type);
  const captions = await getExistingCaptions(context, type);

  const issues = [];
  const seen = new Set();
  captions.forEach((c) => {
    if (c.number !== c.expected) {
      issues.push(`${label} ${c.number} should be ${label} ${c.expected} (click Renumber All).`);
    }
    if (seen.has(c.number)) issues.push(`Duplicate ${label} number ${c.number} found.`);
    seen.add(c.number);
    if (!c.title) issues.push(`${label} ${c.number} is missing caption text.`);
  });

  return { valid: issues.length === 0, count: captions.length, issues };
}
