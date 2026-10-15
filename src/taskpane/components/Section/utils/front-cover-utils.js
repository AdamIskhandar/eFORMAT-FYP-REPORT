/**
 * Front Cover generation - Most-front Cover ONLY (Inside Cover comes later).
 *
 * This page uses a DIFFERENT margin (4cm/4cm/3.2cm/3.2cm) than the rest of
 * the report (set by Title Page's "Set Margin", 3cm/2.5cm/3.8cm/2.5cm). To
 * support that, Front Cover gets its own Word section (section index 0),
 * with its own margin applied directly - and margin.js's applyUniformMargin
 * is told to skip that section via a "protected section count" flag, so
 * Title Page's margin button never overwrites it.
 *
 * Built once (idempotent via a marker bookmark) with bracket placeholder
 * text ("[Write your project title]") that the student edits directly in
 * Word - same approach as the Table of Contents skeleton.
 *
 * Elements, matching the sample cover image:
 *   - UNIVERSITI POLY-TECH MALAYSIA - Arial Bold 20 - First line on page
 *   - Project Title                 - Arial Bold 20 - Centre of Page
 *   - Student's Full Name           - Arial Bold 20 - Last line on page - RED
 *   - Programme line: "BACHELOR OF INFORMATION TECHNOLOGY (HONS) IN" in
 *     black, then the specialisation (e.g. "COMPUTER APPLICATION
 *     DEVELOPMENT") in RED, as its own placeholder since it differs by
 *     programme/intake.
 *
 * Vertical position ("First line", "Centre of Page", "Last line on page")
 * is approximated with spaceBefore, the same technique already used in
 * Title Page / Declaration - not pixel-exact, the student can nudge spacing
 * in Word if needed.
 */

/* global Word, Office */

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
const RED = "#FF0000";

export function isFrontCoverGenerated() {
  return Office.context.document.settings.get(PROTECTED_SECTIONS_KEY) > 0;
}

function markFrontCoverSections(count) {
  Office.context.document.settings.set(PROTECTED_SECTIONS_KEY, count);
  Office.context.document.settings.saveAsync();
}

function formatLine(p, { size, spaceBefore = 0, spaceAfter = 0, color = BLACK }) {
  p.font.name = "Arial";
  p.font.bold = true;
  p.font.size = size;
  p.font.color = color;
  p.alignment = "Centered";
  p.spaceBefore = spaceBefore;
  p.spaceAfter = spaceAfter;
  p.lineSpacing = 12; // "Single" spacing, matches the rest of this project's convention
  return p;
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

function addLine(anchorParagraph, text, opts) {
  return formatLine(anchorParagraph.insertParagraph(text, "After"), opts);
}

export async function generateFrontCover(context) {
  const marker = context.document.getBookmarkRangeOrNullObject(MARKER_BOOKMARK);
  marker.load("isNullObject");
  await context.sync();
  if (!marker.isNullObject) {
    return { created: false };
  }

  // ---------------- Most-front Cover (one page, one section) ----------------
  // The university line IS the first paragraph of the document - no blank
  // leading paragraph - so its own height doesn't push the rest of the
  // content past the bottom margin and spill onto a second page.
  let cursor = formatLine(
    context.document.body.insertParagraph(
      "UNIVERSITI POLY-TECH MALAYSIA",
      Word.InsertLocation.start
    ),
    {
      size: 22,
      spaceBefore: 0,
    }
  );
  // Spacing kept modest (not literal "exact centre"/"exact bottom" math) so
  // all four lines - including the two-line programme text - reliably fit
  // within one page at the 4cm top/bottom margin.
  cursor = addLine(cursor, "PROJECT TITLE", { size: 20, spaceBefore: 220 });
  cursor = addLine(cursor, "STUDENT FULL NAME", { size: 20, spaceBefore: 140, color: RED });

  // Programme line: black lead-in + red specialisation placeholder, matching
  // the sample cover's partial-colour styling.
  cursor = addLine(cursor, "BACHELOR OF INFORMATION TECHNOLOGY (HONS) IN ", {
    size: 20,
    spaceBefore: 100,
  });
  const programmeRange = cursor
    .getRange("End")
    .insertText("COMPUTER APPLICATION DEVELOPMENT", "End");
  programmeRange.font.name = "Arial";
  programmeRange.font.bold = true;
  programmeRange.font.size = 20;
  programmeRange.font.color = RED;

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
