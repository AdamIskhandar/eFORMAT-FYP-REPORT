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

import { numberToBookmark, bookmarkToNumber } from "./word-caption-utils";

/* global Word */

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
  const all = context.document.body.getRange("Whole").getBookmarks(true, true);
  await context.sync();
  let names = all.value;

  if (type) {
    // Caption bookmarks only, in number order: Figure_2_1 < Figure_2_2 < Figure_3_1.
    const parts = (n) => bookmarkToNumber(type, n).split(".").map(Number);
    names = names
      .filter((n) => bookmarkToNumber(type, n) !== null)
      .sort((a, b) => {
        const pa = parts(a);
        const pb = parts(b);
        for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
          const d = (pa[i] || 0) - (pb[i] || 0);
          if (d !== 0) return d;
        }
        return 0;
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
  const { label } = getConfig(type);
  const bookmarkName = numberToBookmark(type, number);

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
 * Pass the numberMap it returns ({ "2": "3.1", ... }). References are the
 * blue "Figure 2" links pointing at #Figure_2; each is changed to the new
 * number and link. All references are found BEFORE any is changed, so
 * swaps (2 -> 3 and 3 -> 2 in one pass) can't be applied twice.
 */
export async function updateReferencesAfterRenumber(context, type, numberMap) {
  const { label } = getConfig(type);
  const body = context.document.body;

  const changed = Object.keys(numberMap).filter((oldNumber) => oldNumber !== numberMap[oldNumber]);
  if (changed.length === 0) return 0;

  const searches = changed.map((oldNumber) => {
    const res = body.search(`${label} ${oldNumber}`, { matchCase: false });
    res.load("items/hyperlink");
    return { oldNumber, res };
  });
  await context.sync();

  const edits = [];
  searches.forEach(({ oldNumber, res }) => {
    const oldLink = `#${numberToBookmark(type, oldNumber)}`;
    res.items.forEach((r) => {
      if (r.hyperlink === oldLink) edits.push({ r, newNumber: numberMap[oldNumber] });
    });
  });

  edits.forEach(({ r, newNumber }) => {
    const updated = r.insertText(`${label} ${newNumber}`, Word.InsertLocation.replace);
    updated.hyperlink = `#${numberToBookmark(type, newNumber)}`;
  });
  await context.sync();
  return edits.length;
}
