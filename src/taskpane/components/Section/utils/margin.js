/* global Office */

const cmToPt = (cm) => cm * 28.346;

export const TARGET_MARGIN = {
  top: cmToPt(3),
  bottom: cmToPt(2.5),
  left: cmToPt(3.8),
  right: cmToPt(2.5),
};

// Persisted in the document itself, so it survives add-in reloads and
// re-opening the file, and works the same way across every page component.
const MARGIN_FLAG_KEY = "docToolkitMarginSet";

export function isMarginFlagSet() {
  return Office.context.document.settings.get(MARGIN_FLAG_KEY) === true;
}

function markMarginFlagSet() {
  Office.context.document.settings.set(MARGIN_FLAG_KEY, true);
  Office.context.document.settings.saveAsync();
}

// Applies the same margin to every section currently in the document, then
// marks the flag so this never needs to run again.
export async function applyUniformMargin(context) {
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();

  sections.items.forEach((section) => {
    section.pageSetup.topMargin = TARGET_MARGIN.top;
    section.pageSetup.bottomMargin = TARGET_MARGIN.bottom;
    section.pageSetup.leftMargin = TARGET_MARGIN.left;
    section.pageSetup.rightMargin = TARGET_MARGIN.right;
  });
  await context.sync();

  markMarginFlagSet();
}
