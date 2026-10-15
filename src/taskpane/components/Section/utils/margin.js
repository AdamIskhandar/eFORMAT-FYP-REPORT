/* global Word */

const inchToPt = (inch) => inch * 72;

// Report margin, applied to SECTION 3 only (Title Page onward).
// Left is wider for binding.
export const TARGET_MARGIN = {
  top: inchToPt(1),
  bottom: inchToPt(1.2),
  left: inchToPt(1.52),
  right: inchToPt(1.2),
};

// Section 3 = index 2. Sections 1 and 2 are the Front Cover (see
// front-cover-utils.js) and keep their own margin.
const REPORT_SECTION_INDEX = 2;

// Word stores margins with small rounding, so compare within 1 point.
const TOLERANCE_PT = 1;

async function getReportSection(context) {
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();
  if (sections.items.length <= REPORT_SECTION_INDEX) return null;
  return sections.items[REPORT_SECTION_INDEX];
}

// Checks the REAL margins of section 3 instead of a saved flag. A saved flag
// isn't reverted by Undo (Ctrl+Z), so after undoing the margin the button
// would wrongly say "The margin was set" and refuse to run again.
export async function isReportMarginSet(context) {
  const section = await getReportSection(context);
  if (!section) return false;

  section.pageSetup.load("topMargin,bottomMargin,leftMargin,rightMargin");
  await context.sync();

  const s = section.pageSetup;
  const close = (a, b) => Math.abs(a - b) < TOLERANCE_PT;
  return (
    close(s.topMargin, TARGET_MARGIN.top) &&
    close(s.bottomMargin, TARGET_MARGIN.bottom) &&
    close(s.leftMargin, TARGET_MARGIN.left) &&
    close(s.rightMargin, TARGET_MARGIN.right)
  );
}

// Sets the report margin on section 3 only. (Name kept so existing imports
// still work.)
export async function applyUniformMargin(context) {
  const section = await getReportSection(context);
  if (!section) {
    throw new Error(
      "Generate the Front Cover first. It creates section 3, where the Title Page margin is set."
    );
  }

  const setup = section.pageSetup;
  setup.paperSize = "A4"; // before the margins, so Word doesn't re-adjust them
  setup.topMargin = TARGET_MARGIN.top;
  setup.bottomMargin = TARGET_MARGIN.bottom;
  setup.leftMargin = TARGET_MARGIN.left;
  setup.rightMargin = TARGET_MARGIN.right;
  await context.sync();
}
