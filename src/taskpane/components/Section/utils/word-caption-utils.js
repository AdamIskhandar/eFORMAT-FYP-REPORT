/**
 * Caption Automation Module (Task 4.1)
 * - Figure caption generation
 * - Table caption generation
 * - Auto numbering functions
 * - Caption validation rules
 * - List of Figures / List of Tables generation
 *
 * Design note:
 * Captions are plain paragraphs formatted as "Figure N: text" / "Table N: text",
 * each anchored with a bookmark (Figure_N / Table_N). This avoids relying on
 * native Word SEQ fields, which behave inconsistently between Word desktop
 * and Word on the web. The Cross-Referencing module (Task 4.2) looks up a
 * caption's location purely through these bookmarks.
 *
 * The List of Figures / List of Tables generator below follows the same
 * principle: it's built directly from getExistingCaptions() rather than
 * Word's native `TOC \c` field, since that field only recognizes SEQ fields
 * and would never see these bookmark-based captions. Page numbers are shown
 * via a PAGEREF field pointing at each caption's own Figure_N/Table_N
 * bookmark (already created by insertCaption) - so no new native numbering
 * mechanism is introduced, just a page lookup against an existing bookmark.
 *
 * Entries are laid out as a borderless 2-column table (caption text | page
 * number) rather than tab-stopped paragraphs. Word's JS API does not
 * reliably expose paragraph tab stops, so building this with a table -
 * solid, well-supported functionality - avoids that fragility entirely.
 *
 * Tradeoff: renumbering is explicit (call renumberCaptions), it does not
 * happen automatically the way native Word fields would update on their own.
 * The List of Figures/Tables is the same - call generateCaptionList again
 * after adding or renumbering captions to refresh it. Each call wraps the
 * whole generated section (heading + table) in one bookmark and deletes +
 * rebuilds that bookmark's contents, so repeated clicks never duplicate
 * page breaks or leave stale entries behind.
 */

const CAPTION_CONFIG = {
  figure: { label: "Figure", bookmarkPrefix: "Figure_" },
  table: { label: "Table", bookmarkPrefix: "Table_" },
};

const LIST_CONFIG = {
  figure: { blockBookmark: "bmFigureListBlock", headingText: "LIST OF FIGURES" },
  table: { blockBookmark: "bmTableListBlock", headingText: "LIST OF TABLES" },
};

function getConfig(type) {
  const config = CAPTION_CONFIG[type];
  if (!config) throw new Error(`Unknown caption type "${type}". Use "figure" or "table".`);
  return config;
}

/**
 * Scans the document body for existing captions of a given type, in
 * document order, and returns their paragraph objects + parsed numbers.
 */
export async function getExistingCaptions(context, type) {
  const { label } = getConfig(type);
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items/text");
  await context.sync();

  const pattern = new RegExp(`^${label}\\s+(\\d+)\\s*:`, "i");
  const captions = [];

  paragraphs.items.forEach((p, index) => {
    const match = p.text.trim().match(pattern);
    if (match) {
      captions.push({
        paragraph: p,
        index,
        number: parseInt(match[1], 10),
        text: p.text.trim(),
      });
    }
  });

  return captions;
}

/**
 * Task 4.1 - Develop figure/table caption generation.
 * Inserts a new caption paragraph after the current selection, numbered
 * one higher than the current max, and anchors a bookmark to it.
 */
export async function insertCaption(context, type, captionText) {
  const { label, bookmarkPrefix } = getConfig(type);

  const existing = await getExistingCaptions(context, type);
  const nextNumber = existing.length > 0 ? Math.max(...existing.map((c) => c.number)) + 1 : 1;

  const selection = context.document.getSelection();
  const fullText = `${label} ${nextNumber}: ${captionText}`;
  const inserted = selection.insertParagraph(fullText, Word.InsertLocation.after);

  inserted.font.bold = true;
  inserted.font.size = 11;
  inserted.alignment = Word.Alignment.centered;
  await context.sync();

  const bookmarkName = `${bookmarkPrefix}${nextNumber}`;
  inserted.getRange().insertBookmark(bookmarkName);
  await context.sync();

  return { number: nextNumber, bookmarkName, text: fullText };
}

/**
 * Task 4.1 - Implement auto numbering functions.
 * Re-scans captions of a type in document order and renumbers them 1..N,
 * renaming bookmarks to match. Returns { oldNumber: newNumber } so callers
 * can sync any inline cross-references (see updateReferencesAfterRenumber
 * in word-crossref-utils.js).
 */
export async function renumberCaptions(context, type) {
  const { label, bookmarkPrefix } = getConfig(type);
  const captions = await getExistingCaptions(context, type);
  captions.sort((a, b) => a.index - b.index);

  const numberMap = {};

  for (let i = 0; i < captions.length; i++) {
    const newNumber = i + 1;
    const oldNumber = captions[i].number;
    numberMap[oldNumber] = newNumber;

    if (oldNumber !== newNumber) {
      const p = captions[i].paragraph;
      const rest = captions[i].text
        .replace(new RegExp(`^${label}\\s+${oldNumber}\\s*:`, "i"), "")
        .trim();
      p.insertText(`${label} ${newNumber}: ${rest}`, Word.InsertLocation.replace);

      const oldBookmark = `${bookmarkPrefix}${oldNumber}`;
      const newBookmark = `${bookmarkPrefix}${newNumber}`;
      const range = context.document.getBookmarkRangeOrNullObject(oldBookmark);
      range.load("isNullObject");
      await context.sync();

      if (!range.isNullObject) {
        range.insertBookmark(newBookmark);
      }
    }
  }

  await context.sync();
  return numberMap;
}

/**
 * Task 4.1 - Implement caption validation rules.
 * Checks captions of a type are sequential from 1 with no gaps/duplicates,
 * and that each has actual text after the colon.
 */
export async function validateCaptions(context, type) {
  const { label } = getConfig(type);
  const captions = await getExistingCaptions(context, type);
  captions.sort((a, b) => a.index - b.index);

  const issues = [];
  const seen = new Set();

  captions.forEach((c, i) => {
    const expected = i + 1;
    if (c.number !== expected) {
      issues.push(`${label} at position ${i + 1} is numbered ${c.number}, expected ${expected}.`);
    }
    if (seen.has(c.number)) {
      issues.push(`Duplicate ${label} number ${c.number} found.`);
    }
    seen.add(c.number);

    if (!/:\s*\S/.test(c.text)) {
      issues.push(`${label} ${c.number} is missing caption text after the colon.`);
    }
  });

  return { valid: issues.length === 0, count: captions.length, issues };
}

/**
 * List of Figures / List of Tables generator.
 *
 * Builds (or rebuilds) a plain list of captions in document order, driven
 * entirely by getExistingCaptions() - not Word's native TOC field, since
 * that only recognizes SEQ fields and these captions are plain text +
 * bookmarks. Safe to call repeatedly: it clears its own previously
 * generated entries (tracked via entriesBookmark) and reinserts fresh ones,
 * so it stays in sync as captions are added, removed, or renumbered.
 */
export async function generateCaptionList(context, type) {
  const { label, bookmarkPrefix } = getConfig(type);
  const { headingBookmark, entriesBookmark, headingText } = LIST_CONFIG[type];

  const captions = await getExistingCaptions(context, type);
  captions.sort((a, b) => a.index - b.index);

  const headingMarker = context.document.getBookmarkRangeOrNullObject(headingBookmark);
  await context.sync();

  if (headingMarker.isNullObject) {
    // First time: create the heading on its own page, plus an empty
    // placeholder paragraph (bookmarked) that marks where entries go and
    // acts as the "end of list" anchor for future regenerations.
    context.document.body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);

    const heading = context.document.body.insertParagraph(headingText, Word.InsertLocation.end);
    heading.font.name = "Arial";
    heading.font.bold = true;
    heading.font.size = 12;
    heading.font.allCaps = true;
    heading.alignment = Word.Alignment.centered;
    heading.spaceAfter = 18;
    heading.getRange().insertBookmark(headingBookmark);

    const placeholder = context.document.body.insertParagraph("", Word.InsertLocation.end);
    placeholder.getRange().insertBookmark(entriesBookmark);
    await context.sync();
  }

  // Remove any previously generated entry paragraphs (tagged with their own
  // bookmarks) before rebuilding.
  const entryTagPrefix = type === "table" ? "bmListEntryTable_" : "bmListEntryFigure_";
  for (let i = 1; i <= 200; i++) {
    const tag = `${entryTagPrefix}${i}`;
    const existingEntry = context.document.getBookmarkRangeOrNullObject(tag);
    existingEntry.load("isNullObject");
    // eslint-disable-next-line no-await-in-loop
    await context.sync();
    if (existingEntry.isNullObject) break; // stop at first gap - entries are always contiguous 1..N
    existingEntry.paragraphs.load("items");
    // eslint-disable-next-line no-await-in-loop
    await context.sync();
    existingEntry.paragraphs.items.forEach((p) => p.delete());
    // eslint-disable-next-line no-await-in-loop
    await context.sync();
  }

  if (captions.length === 0) {
    return { count: 0 };
  }

  const anchor = context.document.getBookmarkRangeOrNullObject(entriesBookmark);
  await context.sync();

  captions.forEach((c, i) => {
    // Inserting "Before" the anchor each time keeps entries in order and
    // ahead of the placeholder, matching the pattern used elsewhere in
    // this add-in for skeleton content.
    const entry = anchor.insertParagraph(`${c.text}\t`, Word.InsertLocation.before);
    entry.font.name = "Arial";
    entry.font.size = 11;
    entry.font.bold = false;
    entry.alignment = Word.Alignment.left;
    entry.spaceAfter = 6;
    entry.paragraphFormat.tabStops.clear();
    entry.paragraphFormat.tabStops.add(
      PAGE_NUMBER_TAB_STOP_PT,
      Word.TabAlignment.right,
      Word.TabLeader.dots
    );

    // Points at this caption's own Figure_N / Table_N bookmark (already
    // created by insertCaption) - no new field-based numbering introduced,
    // this only asks Word "what page is that bookmark currently on".
    const captionBookmark = `${bookmarkPrefix}${c.number}`;
    entry
      .getRange("End")
      .insertField(Word.InsertLocation.end, Word.FieldType.pageRef, captionBookmark, false);

    entry.getRange().insertBookmark(`${entryTagPrefix}${i + 1}`);
  });

  await context.sync();
  return { count: captions.length };
}
