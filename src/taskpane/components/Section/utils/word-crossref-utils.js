/**
 * Cross-Referencing Module (Task 4.2)
 * - Figure referencing functions
 * - Table referencing functions
 * - Bookmark management
 * - Hyperlink navigation
 *
 * Depends on captions created by word-caption-utils.js, since references
 * are resolved against the Figure_N / Table_N bookmarks that module creates.
 */

const REF_CONFIG = {
  figure: { label: "Figure", bookmarkPrefix: "Figure_" },
  table: { label: "Table", bookmarkPrefix: "Table_" },
};

function getConfig(type) {
  const config = REF_CONFIG[type];
  if (!config) throw new Error(`Unknown reference type "${type}". Use "figure" or "table".`);
  return config;
}

/**
 * Task 4.2 - Bookmark management: list all caption bookmarks of a type
 * (or all bookmarks in the document if no type given), sorted by number.
 */
export async function listBookmarks(context, type = null) {
  if (!context.document.bookmarks) {
    throw new Error("Bookmarks collection is not available on this Word host/API version.");
  }

  context.document.bookmarks.load("items");
  await context.sync();

  let names = context.document.bookmarks.items.map((b) => b.name);

  if (type) {
    const { bookmarkPrefix } = getConfig(type);
    names = names
      .filter((n) => n.startsWith(bookmarkPrefix))
      .sort((a, b) => {
        const na = parseInt(a.replace(bookmarkPrefix, ""), 10);
        const nb = parseInt(b.replace(bookmarkPrefix, ""), 10);
        return na - nb;
      });
  }

  return names;
}

/**
 * Task 4.2 - Bookmark management: delete a bookmark by name (removes the
 * anchor only, not the underlying text).
 */
export async function deleteBookmark(context, bookmarkName) {
  const range = context.document.getBookmarkRangeOrNullObject(bookmarkName);
  range.load("isNullObject");
  await context.sync();

  if (range.isNullObject) return false;

  range.delete();
  await context.sync();
  return true;
}

/**
 * Task 4.2 - Develop figure/table referencing functions.
 * Inserts clickable text like "see Figure 2" at the cursor, hyperlinked
 * to the caption's bookmark. Throws if the caption doesn't exist yet.
 */
export async function insertReference(context, type, number, prefixText = "") {
  const { label, bookmarkPrefix } = getConfig(type);
  const bookmarkName = `${bookmarkPrefix}${number}`;

  const target = context.document.getBookmarkRangeOrNullObject(bookmarkName);
  target.load("isNullObject");
  await context.sync();

  if (target.isNullObject) {
    throw new Error(`No caption found for ${label} ${number}. Insert that caption first.`);
  }

  const selection = context.document.getSelection();
  const refText = `${prefixText}${label} ${number}`;
  const inserted = selection.insertText(refText, Word.InsertLocation.replace);

  // Task 4.2 - Develop hyperlink navigation
  // "#bookmarkName" creates an internal link that jumps to the bookmark
  // when the user clicks it in the document.
  inserted.hyperlink = `#${bookmarkName}`;
  inserted.font.underline = Word.UnderlineType.single;
  inserted.font.color = "#1155CC";

  await context.sync();
  return { bookmarkName, text: refText };
}

/**
 * Task 4.2 - Hyperlink navigation helper.
 * Programmatically scrolls to and selects a bookmark's location — useful
 * for a "Go to Figure N" button in the task pane, independent of the
 * clickable in-document hyperlink.
 */
export async function navigateToBookmark(context, bookmarkName) {
  const range = context.document.getBookmarkRangeOrNullObject(bookmarkName);
  range.load("isNullObject");
  await context.sync();

  if (range.isNullObject) {
    throw new Error(`Bookmark "${bookmarkName}" not found.`);
  }

  range.select();
  await context.sync();
  return true;
}

/**
 * Keeps existing inline references in sync after renumberCaptions() runs.
 * Pass the numberMap it returns. Searches for reference runs carrying a
 * hyperlink to the old bookmark and updates their visible text/link.
 * Processes highest-old-number first to avoid double-replacing collisions
 * (e.g. renumbering 2->1 and 1->2 in the same pass).
 */
export async function updateReferencesAfterRenumber(context, type, numberMap) {
  const { label, bookmarkPrefix } = getConfig(type);
  const body = context.document.body;

  const oldNumbers = Object.keys(numberMap)
    .map(Number)
    .sort((a, b) => b - a);

  for (const oldNumber of oldNumbers) {
    const newNumber = numberMap[oldNumber];
    if (oldNumber === newNumber) continue;

    const oldBookmark = `${bookmarkPrefix}${oldNumber}`;
    const newBookmark = `${bookmarkPrefix}${newNumber}`;

    const searchResults = body.search(`${label} ${oldNumber}`, { matchWholeWord: true });
    searchResults.load("items/hyperlink");
    await context.sync();

    searchResults.items.forEach((r) => {
      if (r.hyperlink === `#${oldBookmark}`) {
        r.insertText(`${label} ${newNumber}`, Word.InsertLocation.replace);
        r.hyperlink = `#${newBookmark}`;
      }
    });
    await context.sync();
  }
}
