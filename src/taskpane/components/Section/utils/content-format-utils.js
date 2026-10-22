/**
 * Content formatting (UPTM format table):
 *   Heading                              Arial 16 Bold
 *   Sub heading                          Arial 14 Bold
 *   Content                              Arial 11, 1.5 spacing, justified
 *   Text in a table / Figure-Table label Arial 10
 *
 * formatSelection(context, kind) - formats the paragraphs the cursor or
 *   selection touches as the chosen kind.
 * formatWholeReport(context) - goes through the report from Chapter 1 to
 *   the end and formats every content paragraph, every paragraph inside a
 *   table, and every "Figure 3.1 ..." / "Table 3.1 ..." caption. Headings
 *   are left alone (the Table of Contents sets the chapter heading format),
 *   and so are the front pages before Chapter 1.
 *
 * Only the font name, size, spacing and alignment are set on content, so a
 * word the student made bold or italic stays that way.
 */

/* global Word */

// Change these to change the format.
export const CONTENT_FORMAT = {
  heading: { label: "Heading", font: "Arial", size: 16, bold: true },
  subheading: { label: "Sub heading", font: "Arial", size: 14, bold: true },
  content: { label: "Content", font: "Arial", size: 11, lineSpacing: 18, alignment: "Justified" },
  small: { label: "Text in table / Label Figure or Table", font: "Arial", size: 10, lineSpacing: 12 },
};
// lineSpacing is in points: 18pt = 1.5 lines (as used across this add-in);
// 12pt = single lines, kept for table text so tables stay compact.

// A caption: "Figure 3.1: ..." (":" "." or "-" after the number), or a
// centered "Figure 3.1 ..." line. "Figure 3.1 shows ..." is a sentence.
const CAPTION_PATTERN = /^(Figure|Table)\s+\d+(?:\.\d+)*\s*([:.\-\u2013]?)/i;
const isCaption = (p, text) => {
  const m = text.match(CAPTION_PATTERN);
  return !!m && (m[2] !== "" || p.alignment === "Centered");
};
const PLACEHOLDER_GREY = "#8A8A8A"; // colour of the old hint lines
const CHAPTER_PATTERN = /^CHAPTER\s+(\d+|[IVXLCDM]+)\b/i;

const isHeading = (p) =>
  /^Heading\d$/i.test(p.styleBuiltIn || "") || /^Heading\s*\d/i.test(p.style || "");

function applyKind(paragraph, kind) {
  const f = CONTENT_FORMAT[kind];
  if (!f) throw new Error(`Unknown format "${kind}".`);
  paragraph.font.name = f.font;
  paragraph.font.size = f.size;
  if (f.bold !== undefined) paragraph.font.bold = f.bold;
  if (f.lineSpacing) paragraph.lineSpacing = f.lineSpacing;
  if (f.alignment) paragraph.alignment = f.alignment;
}

/** Formats the paragraphs in the current selection as one kind. */
export async function formatSelection(context, kind) {
  const paragraphs = context.document.getSelection().paragraphs;
  paragraphs.load("items");
  await context.sync();
  if (paragraphs.items.length === 0) throw new Error("Put the cursor in the text you want to format.");
  paragraphs.items.forEach((p) => applyKind(p, kind));
  await context.sync();
  return paragraphs.items.length;
}

/** Formats the whole report, Chapter 1 to the end. */
export async function formatWholeReport(context) {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items/text,items/style,items/styleBuiltIn,items/tableNestingLevel,items/alignment,items/font/color");
  await context.sync();

  // Start at the first chapter heading, so the cover, Title Page,
  // Declaration, Abstract, Contents and lists keep their own formats.
  const start = paragraphs.items.findIndex(
    (p) => p.tableNestingLevel === 0 && isHeading(p) && CHAPTER_PATTERN.test(p.text.trim())
  );
  if (start < 0) {
    throw new Error('No chapter heading found. Generate the Table of Contents first (it adds "CHAPTER 1" ...).');
  }

  const counts = { content: 0, small: 0 };
  paragraphs.items.slice(start).forEach((p) => {
    if (isHeading(p)) return;
    const text = p.text.trim();
    if (p.tableNestingLevel > 0 || isCaption(p, text)) {
      applyKind(p, "small");
      counts.small++;
    } else if (text) {
      applyKind(p, "content");
      // Text typed over an old grey "Write your content here." hint keeps
      // the hint's grey italic look - make it normal black text.
      if ((p.font.color || "").toUpperCase() === PLACEHOLDER_GREY) {
        p.font.color = "#000000";
        p.font.italic = false;
      }
      counts.content++;
    }
  });
  await context.sync();
  return counts;
}
