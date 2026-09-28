/**
 * Chapter-based caption numbering + List of Figures / List of Tables
 *
 * updateCaptionList(context, "figure" | "table"):
 * 1. Walks the document in order, tracking the current Heading's number
 *    ("CHAPTER 3: ..." -> "3", "3.3 Requirement Analysis" -> "3.3").
 * 2. Every centered paragraph that starts with "Figure ..." / "Table ..." is
 *    a caption. It's renumbered to <heading number>.<count in that heading>,
 *    e.g. Figure 3.1, Figure 3.3.1, Figure 3.3.2. Only the label is replaced,
 *    so the caption's own formatting is kept.
 * 3. Captions in chapters before FIRST_CAPTION_CHAPTER are ignored.
 * 4. Each caption is bookmarked, and the list is rebuilt IN PLACE under its
 *    heading (old table deleted, new one inserted) with PAGEREF page numbers.
 *
 * The list heading is created once: right before Chapter 1 if the Table of
 * Contents was generated (so it sits in the front matter), otherwise at the
 * end of the document.
 */

/* global Word */

export const FIRST_CAPTION_CHAPTER = 1;

const FIRST_HEADING_BOOKMARK = "bmReportHeading_0"; // created by the TOC generator

const TYPE_CONFIG = {
  figure: {
    label: "Figure",
    headingText: "LIST OF FIGURES",
    headingBookmark: "bmFigureListHeading",
    captionBookmarkPrefix: "bmFig_",
  },
  table: {
    label: "Table",
    headingText: "LIST OF TABLES",
    headingBookmark: "bmTableListHeading",
    captionBookmarkPrefix: "bmTbl_",
  },
};

// Usable text width on A4 with 3.8cm / 2.5cm side margins, in points.
const PAGE_COL_WIDTH_PT = 46;
const TEXT_COL_WIDTH_PT = 370;

function headingPrefix(text) {
  const t = text.trim();
  let m = t.match(/^CHAPTER\s+(\d+)\b/i);
  if (m) return m[1];
  // "3 METHODOLOGY", "3.3 Phases in Agile Methodology", "3.5.1.1 Questionnaire ..."
  m = t.match(/^(\d+(?:\.\d+)*)\.?\s+\S/);
  if (m) return m[1];
  return null;
}

async function scanAndRenumber(context, cfg) {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load(
    "items/text,items/style,items/styleBuiltIn,items/alignment,items/isListItem,items/tableNestingLevel,items/font/bold"
  );
  await context.sync();

  // A heading is a Heading 1-9 styled paragraph, OR - for subheadings a student
  // typed by hand - a short, bold line that starts with a section number
  // ("3.1 Introduction", "CHAPTER 3: ...") and doesn't read like a sentence.
  const isHeadingPara = (p) => {
    if (/^Heading\s*\d$/i.test(p.styleBuiltIn || "") || /^Heading\s*\d/i.test(p.style || ""))
      return true;
    const t = p.text.trim();
    return (
      p.font.bold === true &&
      t.length > 0 &&
      t.length <= 120 &&
      !/[.;]$/.test(t) &&
      headingPrefix(t) !== null
    );
  };

  // Heading numbers come from the heading text ("3.3 Requirement Analysis"),
  // or - if the template uses automatic list numbering, so the number isn't
  // in the text - from the paragraph's list string.
  const headingPrefixByIndex = new Map();
  const pendingLists = [];
  paragraphs.items.forEach((p, idx) => {
    if (p.tableNestingLevel > 0 || !isHeadingPara(p)) return;
    const fromText = headingPrefix(p.text);
    if (fromText) {
      headingPrefixByIndex.set(idx, fromText);
    } else if (p.isListItem) {
      const li = p.listItemOrNullObject;
      li.load("isNullObject,listString");
      pendingLists.push({ idx, li });
    }
  });
  await context.sync();
  pendingLists.forEach(({ idx, li }) => {
    if (li.isNullObject) return;
    const m = li.listString.match(/(\d+(?:\.\d+)*)/);
    if (m) headingPrefixByIndex.set(idx, m[1]);
  });

  const captionPattern = new RegExp(
    `^(${cfg.label}\\s+\\d+(?:\\.\\d+)*\\s*[:.\\-\\u2013]?)\\s*(.*)$`,
    "i"
  );

  const diagnostics = {
    headings: headingPrefixByIndex.size,
    captionLike: 0,
    notCentered: 0,
    noHeadingAbove: 0,
    beforeFirstChapter: 0,
  };
  let currentPrefix = null;
  const counters = {};
  const found = [];

  paragraphs.items.forEach((p, idx) => {
    if (p.tableNestingLevel > 0) return; // ignore the generated lists themselves
    if (isHeadingPara(p)) {
      // A heading with no number ("Appendix A - ...", "References") ends the
      // numbered chapters, so figures after it must not inherit the last
      // chapter's number.
      currentPrefix = headingPrefixByIndex.has(idx) ? headingPrefixByIndex.get(idx) : null;
      return;
    }

    const m = p.text.trim().match(captionPattern);
    if (!m) return;
    if (p.alignment !== "Centered") {
      diagnostics.notCentered++;
      return;
    }
    diagnostics.captionLike++;

    if (!currentPrefix) {
      diagnostics.noHeadingAbove++;
      return;
    }
    const chapter = parseInt(currentPrefix.split(".")[0], 10);
    if (chapter < FIRST_CAPTION_CHAPTER) {
      diagnostics.beforeFirstChapter++;
      return;
    }

    counters[currentPrefix] = (counters[currentPrefix] || 0) + 1;
    found.push({
      paragraph: p,
      oldPrefix: m[1].trim(),
      number: `${currentPrefix}.${counters[currentPrefix]}`,
      title: m[2].trim(),
    });
  });

  // Replace only the label part of each caption ("Figure 1:" -> "Figure 3.3.2").
  const searches = found.map((c) => {
    const newPrefix = `${cfg.label} ${c.number}`;
    c.entryText = c.title ? `${newPrefix} ${c.title}` : newPrefix;
    if (c.oldPrefix === newPrefix) return null;
    const res = c.paragraph.search(c.oldPrefix, { matchCase: true });
    res.load("items");
    return { res, newPrefix };
  });
  await context.sync();

  searches.forEach((s) => {
    if (s && s.res.items.length > 0) s.res.items[0].insertText(s.newPrefix, "Replace");
  });

  // Bookmark each caption so the list's PAGEREF fields have a target.
  found.forEach((c, i) => {
    c.paragraph.getRange().insertBookmark(`${cfg.captionBookmarkPrefix}${i + 1}`);
  });
  await context.sync();

  return { found, diagnostics };
}

async function getOrCreateListHeading(context, cfg) {
  const existing = context.document.getBookmarkRangeOrNullObject(cfg.headingBookmark);
  existing.load("isNullObject");
  await context.sync();
  if (!existing.isNullObject) return existing.paragraphs.getFirst();

  const anchor = context.document.getBookmarkRangeOrNullObject(FIRST_HEADING_BOOKMARK);
  anchor.load("isNullObject");
  await context.sync();

  let heading;
  if (!anchor.isNullObject) {
    // Front-matter placement: just before Chapter 1, then a page break after
    // the list so Chapter 1 still starts on its own page.
    heading = anchor.paragraphs.getFirst().insertParagraph(cfg.headingText, "Before");
    heading.style = "Normal";
    const spacer = heading.insertParagraph("", "After");
    spacer.style = "Normal";
    spacer.insertBreak(Word.BreakType.page, "After");
  } else {
    context.document.body.insertBreak(Word.BreakType.page, "End");
    heading = context.document.body.insertParagraph(cfg.headingText, "End");
    heading.style = "Normal";
  }

  heading.font.name = "Arial";
  heading.font.bold = true;
  heading.font.size = 12;
  heading.font.allCaps = true;
  heading.alignment = "Centered";
  heading.spaceAfter = 18;
  heading.getRange().insertBookmark(cfg.headingBookmark);
  await context.sync();

  return heading;
}

export async function updateCaptionList(context, type) {
  const cfg = TYPE_CONFIG[type];
  if (!cfg) throw new Error(`Unknown caption type "${type}".`);

  const { found, diagnostics } = await scanAndRenumber(context, cfg);
  const headingPara = await getOrCreateListHeading(context, cfg);

  // Remove the previous table (if the paragraph after the heading is in one).
  const next = headingPara.getNextOrNullObject();
  next.load("isNullObject");
  await context.sync();
  if (!next.isNullObject) {
    const oldTable = next.parentTableOrNullObject;
    oldTable.load("isNullObject");
    await context.sync();
    if (!oldTable.isNullObject) oldTable.delete();
  }

  const values =
    found.length > 0
      ? found.map((c) => [c.entryText, ""])
      : [[`No ${cfg.label.toLowerCase()} captions yet`, ""]];
  const table = headingPara.insertTable(values.length, 2, "After", values);
  table.font.name = "Arial";
  table.font.size = 11;
  table.font.bold = false;

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

  table.getCell(0, 0).columnWidth = TEXT_COL_WIDTH_PT;
  table.getCell(0, 1).columnWidth = PAGE_COL_WIDTH_PT;
  for (let i = 0; i < values.length; i++) {
    table.getCell(i, 1).horizontalAlignment = "Right";
  }
  await context.sync();

  // Page numbers last: the caption bookmarks already exist, so no
  // "Error! Bookmark not defined."
  found.forEach((c, i) => {
    table
      .getCell(i, 1)
      .body.getRange("End")
      .insertField("End", Word.FieldType.pageRef, `${cfg.captionBookmarkPrefix}${i + 1}`, false);
  });
  await context.sync();

  return { count: found.length, diagnostics };
}
