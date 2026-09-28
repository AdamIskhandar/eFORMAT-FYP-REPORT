/**
 * Table of Contents Generation
 *
 * One-time setup action (same idempotency idea as Title Page / Declaration):
 * this creates real content students type into, so a second click must NOT
 * rebuild it. Once the marker bookmark exists, Generate just says so.
 *
 * Built from REPORT_STRUCTURE (toc-template-config.js):
 * 1. A "TABLE OF CONTENTS" page: one table row per entry (text | page number)
 *    with the page number as a PAGEREF field pointing at that entry's heading.
 * 2. The real headings, formatted per level:
 *      Heading 1 - centered,           Arial 18
 *      Heading 2 - left aligned,       Arial 14
 *      Heading 3 - indented,           Arial 12
 *      Heading 4 - indented deeper,    Arial 12 italic (e.g. 3.5.1.1)
 *    with a short grey placeholder line under headings the student writes in.
 */

/* global Word */

const MARKER_BOOKMARK = "bmTocGenerated";
const HEADING_BOOKMARK_PREFIX = "bmReportHeading_";

// Formatting of the real headings in the document body, by level.
const HEADING_FORMAT = {
  1: {
    builtIn: "heading1",
    size: 18,
    bold: true,
    italic: false,
    alignment: "Centered",
    leftIndent: 0,
    spaceBefore: 24,
    spaceAfter: 12,
  },
  2: {
    builtIn: "heading2",
    size: 14,
    bold: true,
    italic: false,
    alignment: "Left",
    leftIndent: 0,
    spaceBefore: 14,
    spaceAfter: 6,
  },
  3: {
    builtIn: "heading3",
    size: 12,
    bold: false,
    italic: false,
    alignment: "Left",
    leftIndent: 36,
    spaceBefore: 10,
    spaceAfter: 6,
  },
  4: {
    builtIn: "heading4",
    size: 12,
    bold: false,
    italic: true,
    alignment: "Left",
    leftIndent: 72,
    spaceBefore: 8,
    spaceAfter: 6,
  },
};

// Formatting of the rows on the TOC page, by level (matches the template look).
const TOC_ENTRY_FORMAT = {
  1: { bold: true, italic: false, indent: 0, spaceBefore: 8 },
  2: { bold: true, italic: false, indent: 14, spaceBefore: 6 },
  3: { bold: false, italic: false, indent: 32, spaceBefore: 0 },
  4: { bold: false, italic: true, indent: 52, spaceBefore: 0 },
};

// Usable width on A4 with 3.8cm / 2.5cm side margins is ~416pt.
const TOC_TEXT_COL_WIDTH_PT = 370;
const TOC_PAGE_COL_WIDTH_PT = 46;

function formatHeading(paragraph, level) {
  const f = HEADING_FORMAT[level] || HEADING_FORMAT[1];
  // Style first, then direct formatting on top (Word's default heading
  // styles are blue and use other fonts/sizes).
  paragraph.styleBuiltIn = Word.BuiltInStyleName[f.builtIn];
  paragraph.font.name = "Arial";
  paragraph.font.size = f.size;
  paragraph.font.bold = f.bold;
  paragraph.font.italic = f.italic;
  paragraph.font.allCaps = false;
  paragraph.font.color = "#000000";
  paragraph.alignment = f.alignment;
  paragraph.leftIndent = f.leftIndent;
  paragraph.spaceBefore = f.spaceBefore;
  paragraph.spaceAfter = f.spaceAfter;
}

function formatPlaceholder(paragraph, level) {
  const f = HEADING_FORMAT[level] || HEADING_FORMAT[1];
  // New paragraphs inherit the previous one's style; force plain body text so
  // a placeholder is never mistaken for a heading.
  paragraph.styleBuiltIn = Word.BuiltInStyleName.normal;
  paragraph.font.name = "Arial";
  paragraph.font.size = 11;
  paragraph.font.bold = false;
  paragraph.font.italic = true;
  paragraph.font.allCaps = false;
  paragraph.font.color = "#8a8a8a";
  paragraph.alignment = "Left";
  paragraph.leftIndent = f.leftIndent;
  paragraph.spaceBefore = 0;
  paragraph.spaceAfter = 6;
}

/**
 * @param {Word.RequestContext} context
 * @param {Array} structure - REPORT_STRUCTURE from toc-template-config.js
 */
export async function generateTableOfContents(context, structure) {
  const marker = context.document.getBookmarkRangeOrNullObject(MARKER_BOOKMARK);
  marker.load("isNullObject");
  await context.sync();

  if (!marker.isNullObject) {
    return { created: false, count: structure.length };
  }

  // --- TOC page ---
  context.document.body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

  const heading = context.document.body.insertParagraph(
    "TABLE OF CONTENTS",
    Word.InsertLocation.end
  );
  heading.styleBuiltIn = Word.BuiltInStyleName.normal;
  heading.font.name = "Arial";
  heading.font.bold = true;
  heading.font.italic = false;
  heading.font.size = 14;
  heading.font.allCaps = true;
  heading.font.color = "#000000";
  heading.alignment = "Centered";
  heading.leftIndent = 0;
  heading.spaceBefore = 0;
  heading.spaceAfter = 18;

  const values = structure.map((entry) => [entry.text, ""]);
  const table = context.document.body.insertTable(
    structure.length,
    2,
    Word.InsertLocation.end,
    values
  );
  table.font.name = "Arial";
  table.font.size = 11;

  [
    Word.BorderLocation.top,
    Word.BorderLocation.bottom,
    Word.BorderLocation.left,
    Word.BorderLocation.right,
    Word.BorderLocation.insideHorizontal,
    Word.BorderLocation.insideVertical,
  ].forEach((loc) => {
    table.getBorder(loc).type = Word.BorderType.none;
  });

  table.getCell(0, 0).columnWidth = TOC_TEXT_COL_WIDTH_PT;
  table.getCell(0, 1).columnWidth = TOC_PAGE_COL_WIDTH_PT;

  await context.sync();

  // Row styling only for now. The PAGEREF fields go in AFTER the headings and
  // their bookmarks exist: Word evaluates a field the moment it's inserted, so
  // a field pointing at a missing bookmark stores "Error! Bookmark not defined."
  structure.forEach((entry, i) => {
    const fmt = TOC_ENTRY_FORMAT[entry.level] || TOC_ENTRY_FORMAT[1];

    const textPara = table.getCell(i, 0).body.paragraphs.getFirst();
    textPara.leftIndent = fmt.indent;
    textPara.spaceBefore = fmt.spaceBefore;
    textPara.spaceAfter = 0;
    textPara.font.bold = fmt.bold;
    textPara.font.italic = fmt.italic;

    const pageCell = table.getCell(i, 1);
    pageCell.horizontalAlignment = Word.Alignment.right;
    const pagePara = pageCell.body.paragraphs.getFirst();
    pagePara.spaceBefore = fmt.spaceBefore;
    pagePara.spaceAfter = 0;
    pagePara.font.bold = fmt.bold;
    pagePara.font.italic = fmt.italic;
  });

  await context.sync();

  // --- The real headings (and placeholders) students write under ---
  context.document.body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

  structure.forEach((entry, i) => {
    const headingPara = context.document.body.insertParagraph(entry.text, Word.InsertLocation.end);
    formatHeading(headingPara, entry.level);
    headingPara.getRange().insertBookmark(`${HEADING_BOOKMARK_PREFIX}${i}`);

    if (entry.placeholder) {
      const placeholderPara = context.document.body.insertParagraph(
        entry.placeholder,
        Word.InsertLocation.end
      );
      formatPlaceholder(placeholderPara, entry.level);
    }
  });

  const trailing = context.document.body.insertParagraph("", Word.InsertLocation.end);
  trailing.styleBuiltIn = Word.BuiltInStyleName.normal;
  trailing.leftIndent = 0;

  // Sync so every heading bookmark really exists...
  await context.sync();

  // ...and only now insert the page-number fields that point at them.
  structure.forEach((entry, i) => {
    table
      .getCell(i, 1)
      .body.getRange("End")
      .insertField(
        Word.InsertLocation.end,
        Word.FieldType.pageRef,
        `${HEADING_BOOKMARK_PREFIX}${i}`,
        false
      );
  });

  await context.sync();

  // One bookmark around the whole generated block, only so a later click can
  // tell it has already been generated.
  const wholeBlockRange = heading.getRange("Start").expandTo(trailing.getRange("End"));
  wholeBlockRange.insertBookmark(MARKER_BOOKMARK);
  await context.sync();

  return { created: true, count: structure.length };
}
