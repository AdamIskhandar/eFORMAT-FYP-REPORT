/**
 * List of Figures / List of Tables
 *
 * updateCaptionList(context, "figure" | "table"):
 * 1. Renumbers every caption by chapter, using the same rule as Insert
 *    Caption (word-caption-utils.js): the first figure in Chapter 3 is
 *    Figure 3.1, the next Figure 3.2. Cross-references follow the new
 *    numbers. Captions are found in any paragraph, centered or not, and
 *    inside tables.
 * 2. Rebuilds the list IN PLACE between its heading and an (invisible) end
 *    line: old lines deleted, new ones inserted. Each entry is ONE line
 *    with tab stops - number | title | page - not a Word table, so there
 *    are no dotted table gridlines. Page numbers are PAGEREF fields that
 *    point at each caption's own bookmark (Figure_3_1 ...), and each entry
 *    links to its caption (Ctrl+Click).
 *
 * The list heading is created once: right before Chapter 1 if the Table of
 * Contents was generated (so it sits in the front matter), otherwise at the
 * end of the document.
 */

/* global Word */

import { renumberCaptions, getExistingCaptions, numberToBookmark } from "./word-caption-utils";
import { updateReferencesAfterRenumber } from "./word-crossref-utils";

export const FIRST_CAPTION_CHAPTER = 1;

const FIRST_HEADING_BOOKMARK = "bmReportHeading_0"; // created by the TOC generator

const TYPE_CONFIG = {
  figure: {
    label: "Figure",
    headingText: "LIST OF FIGURES",
    headingBookmark: "bmFigureListHeading",
    endBookmark: "bmFigureListEnd",
  },
  table: {
    label: "Table",
    headingText: "LIST OF TABLES",
    headingBookmark: "bmTableListHeading",
    endBookmark: "bmTableListEnd",
  },
};

// List layout (UPTM sample):
//   Table No.                                            Page
//   Table 2.1   The fatty acid composition of adipose
//               tissues from various sites ...             41
// Each entry is one paragraph: the number, a tab to the title column, the
// title (wrapped lines line up under it), then a tab to the page number at
// the right margin - so the page number sits on the title's LAST line.
// Change these to move things.
const inchToPt = (inch) => inch * 72;
const TEXT_WIDTH_PT = 595.3 - inchToPt(1.52) - inchToPt(1.2); // A4, section 3 margins
const NUMBER_COL_WIDTH_PT = 86; // where the title starts ("Table 2.1" fits before it)
const PAGE_COL_WIDTH_PT = 44; // space kept free for the page number
const LIST_FONT_SIZE = 12;
const GAP_AFTER_HEADER_PT = 18; // gap under "Table No. / Page"
const GAP_BETWEEN_ENTRIES_PT = 14; // blank line between entries

const toTwips = (pt) => Math.round(pt * 20);

function escapeXml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// caps off on purpose: the list sits under an ALL CAPS heading and would
// otherwise inherit it.
function rPr(bold) {
  const hp = LIST_FONT_SIZE * 2;
  return (
    `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>` +
    (bold ? "<w:b/><w:bCs/>" : '<w:b w:val="0"/><w:bCs w:val="0"/>') +
    `<w:i w:val="0"/><w:caps w:val="0"/><w:color w:val="000000"/><w:u w:val="none"/>` +
    `<w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/></w:rPr>`
  );
}
const textRun = (text, bold = false) =>
  `<w:r>${rPr(bold)}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
const tabRun = (bold = false) => `<w:r>${rPr(bold)}<w:tab/></w:r>`;

const PAGE_TAB = `<w:tab w:val="right" w:pos="${toTwips(TEXT_WIDTH_PT)}"/>`;

function lineXml(runs, { before = 0, after = 0, hanging = false, anchor = null } = {}) {
  const ind = hanging
    ? `<w:ind w:left="${toTwips(NUMBER_COL_WIDTH_PT)}" w:hanging="${toTwips(NUMBER_COL_WIDTH_PT)}" w:right="${toTwips(PAGE_COL_WIDTH_PT)}"/>`
    : `<w:ind w:left="0" w:firstLine="0" w:right="0"/>`;
  const body = anchor
    ? `<w:hyperlink w:anchor="${escapeXml(anchor)}" w:history="1">${runs}</w:hyperlink>`
    : runs;
  return (
    `<w:p><w:pPr><w:pStyle w:val="Normal"/><w:keepNext w:val="0"/><w:tabs>${PAGE_TAB}</w:tabs>` +
    `<w:spacing w:before="${toTwips(before)}" w:after="${toTwips(after)}" w:line="240" w:lineRule="auto"/>` +
    `${ind}<w:jc w:val="left"/>${rPr(false)}</w:pPr>${body}</w:p>`
  );
}

function listOoxml(cfg, found) {
  const header = lineXml(textRun(`${cfg.label} No.`, true) + tabRun(true) + textRun("Page", true), {
    after: GAP_AFTER_HEADER_PT,
  });
  const entries =
    found.length > 0
      ? found
          .map((c, i) =>
            lineXml(textRun(c.label) + tabRun() + textRun(c.title) + tabRun(), {
              before: i === 0 ? 0 : GAP_BETWEEN_ENTRIES_PT,
              hanging: true,
              anchor: c.bookmark,
            })
          )
          .join("")
      : lineXml(textRun(`No ${cfg.label.toLowerCase()} captions yet.`));
  // Throwaway last paragraph: Word merges the LAST inserted paragraph into
  // the line it's inserted before (losing its settings).
  return (
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships></pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${header}${entries}<w:p/></w:body></w:document>` +
    `</pkg:xmlData></pkg:part></pkg:package>`
  );
}

async function collectCaptions(context, type, cfg) {
  const numberMap = await renumberCaptions(context, type);
  await updateReferencesAfterRenumber(context, type, numberMap);

  const captions = await getExistingCaptions(context, type);
  const found = captions.map((c) => ({
    label: `${cfg.label} ${c.expected}`,
    title: c.title,
    bookmark: numberToBookmark(type, c.expected),
  }));
  const chapters = new Set(captions.filter((c) => c.chapter).map((c) => c.chapter));
  const diagnostics = {
    headings: chapters.size,
    captionLike: captions.length,
    notCentered: 0,
    noHeadingAbove: captions.filter((c) => !c.chapter).length,
    beforeFirstChapter: 0,
  };
  return { found, diagnostics };
}

function formatEndLine(paragraph) {
  // Invisible marker line: where the list ends. Tiny, no spacing.
  paragraph.styleBuiltIn = Word.BuiltInStyleName.normal;
  paragraph.font.size = 2;
  paragraph.spaceBefore = 0;
  paragraph.spaceAfter = 0;
  paragraph.lineSpacing = 2;
}

/**
 * Finds (or creates) the list's heading and its end line. A heading
 * bookmark that sits on some other text (left over from an older version)
 * is dropped and a fresh heading made, so a list never lands under the
 * wrong title.
 */
async function getOrCreateListBlock(context, cfg) {
  const existing = context.document.getBookmarkRangeOrNullObject(cfg.headingBookmark);
  existing.load("isNullObject");
  await context.sync();

  let heading = null;
  if (!existing.isNullObject) {
    // Find this list's heading WORDS anywhere in the bookmark: Word can
    // stretch the bookmark over a list inserted in front of it, so its
    // first paragraph may be the OTHER list's heading.
    const hits = existing.search(cfg.headingText, { matchCase: false });
    hits.load("items");
    await context.sync();
    if (hits.items.length > 0) {
      heading = hits.items[0].paragraphs.getFirst();
      hits.items[0].insertBookmark(cfg.headingBookmark); // re-pin on the words
      await context.sync();
    } else {
      context.document.deleteBookmark(cfg.headingBookmark);
      await context.sync();
    }
  }

  if (heading) {
    const end = context.document.getBookmarkRangeOrNullObject(cfg.endBookmark);
    end.load("isNullObject");
    await context.sync();
    if (!end.isNullObject) return { heading, endLine: end.paragraphs.getLast() };

    // Heading from an older version (list was a table): remove that table
    // and add the end line right under the heading.
    const next = heading.getNextOrNullObject();
    next.load("isNullObject");
    await context.sync();
    if (!next.isNullObject) {
      const oldTable = next.parentTableOrNullObject;
      oldTable.load("isNullObject");
      await context.sync();
      if (!oldTable.isNullObject) oldTable.delete();
    }
    const endLine = heading.insertParagraph("", "After");
    formatEndLine(endLine);
    endLine.getRange("Whole").insertBookmark(cfg.endBookmark);
    await context.sync();
    return { heading, endLine };
  }

  const anchor = context.document.getBookmarkRangeOrNullObject(FIRST_HEADING_BOOKMARK);
  anchor.load("isNullObject");
  await context.sync();

  if (!anchor.isNullObject) {
    // Front-matter placement: just before Chapter 1, then a page break after
    // the list so Chapter 1 still starts on its own page.
    heading = anchor.paragraphs.getFirst().insertParagraph(cfg.headingText, "Before");
  } else {
    context.document.body.insertBreak(Word.BreakType.page, "End");
    heading = context.document.body.insertParagraph(cfg.headingText, "End");
  }
  heading.styleBuiltIn = Word.BuiltInStyleName.normal;
  heading.font.name = "Arial";
  heading.font.bold = true;
  heading.font.italic = false;
  heading.font.size = 12;
  heading.font.allCaps = true;
  heading.font.color = "#000000";
  heading.alignment = "Centered";
  heading.leftIndent = 0;
  heading.firstLineIndent = 0;
  heading.spaceBefore = 0;
  heading.spaceAfter = 18;
  // Bookmark the heading WORDS only (never a page break in front of them),
  // so the Contents shows the right page for this list.
  const words = heading.search(cfg.headingText, { matchCase: false });
  words.load("items");
  await context.sync();
  (words.items[0] || heading.getRange("Content")).insertBookmark(cfg.headingBookmark);

  const endLine = heading.insertParagraph("", "After");
  formatEndLine(endLine);
  endLine.font.allCaps = false;
  endLine.font.bold = false;
  endLine.alignment = "Left";
  endLine.getRange("Whole").insertBookmark(cfg.endBookmark);
  if (!anchor.isNullObject) endLine.insertBreak(Word.BreakType.page, "After");
  await context.sync();

  return { heading, endLine };
}

// Heading ... end line, as one range.
function listBlock(heading, endLine) {
  return heading.getRange("Whole").expandTo(endLine.getRange("Whole"));
}

// Deletes the lines between the heading and the end line. Fields go first,
// whole - Word refuses to delete a line that holds part of a field.
async function clearList(context, heading, endLine) {
  const fields = listBlock(heading, endLine).fields;
  fields.load("items");
  await context.sync();
  fields.items.forEach((f) => f.delete());
  await context.sync();

  const lines = listBlock(heading, endLine).paragraphs;
  lines.load("items");
  await context.sync();
  for (let i = lines.items.length - 2; i >= 1; i--) lines.items[i].delete();
  await context.sync();
}

export async function updateCaptionList(context, type) {
  const cfg = TYPE_CONFIG[type];
  if (!cfg) throw new Error(`Unknown caption type "${type}".`);

  const { found, diagnostics } = await collectCaptions(context, type, cfg);
  const { heading, endLine } = await getOrCreateListBlock(context, cfg);

  await clearList(context, heading, endLine);
  endLine.getRange("Start").insertOoxml(listOoxml(cfg, found), "Before");
  await context.sync();

  // Pin the end bookmark back onto the end line only (Word can stretch it
  // over lines inserted right before it).
  const end = context.document.getBookmarkRangeOrNullObject(cfg.endBookmark).paragraphs.getLast();
  end.getRange("Whole").insertBookmark(cfg.endBookmark);
  formatEndLine(end);
  await context.sync();

  // Page numbers last, at the end of each entry line. Lines in the block:
  // [heading, "Table No. / Page", ...entries, end line].
  const lines = listBlock(heading, end).paragraphs;
  lines.load("items");
  await context.sync();
  const FIRST = 2;
  const fields = [];
  found.forEach((c, i) => {
    const line = lines.items[FIRST + i];
    if (!line) return;
    fields.push(line.getRange("End").insertField("End", Word.FieldType.pageRef, c.bookmark, false));
  });
  await context.sync();
  try {
    fields.forEach((f) => f.updateResult());
    await context.sync();
  } catch (err) {
    console.warn(
      "Could not refresh the list page numbers (press F9 in Word if one looks off).",
      err
    );
  }

  return { count: found.length, diagnostics };
}

/**
 * Re-reads the page numbers in both lists. Word lays out the pages in the
 * background AFTER the add-in's changes, so the numbers read straight away
 * can be too early; the panel calls this again a moment later (like F9).
 */
export async function refreshCaptionListPageNumbers(context) {
  let refreshed = 0;
  for (const cfg of Object.values(TYPE_CONFIG)) {
    const head = context.document.getBookmarkRangeOrNullObject(cfg.headingBookmark);
    const end = context.document.getBookmarkRangeOrNullObject(cfg.endBookmark);
    head.load("isNullObject");
    end.load("isNullObject");
    await context.sync();
    if (head.isNullObject || end.isNullObject) continue;

    const fields = listBlock(head.paragraphs.getFirst(), end.paragraphs.getLast()).fields;
    fields.load("items/code");
    await context.sync();
    fields.items
      .filter((f) => /^\s*PAGEREF\b/i.test(f.code || ""))
      .forEach((f) => {
        f.updateResult();
        refreshed++;
      });
    await context.sync();
  }
  return refreshed;
}
