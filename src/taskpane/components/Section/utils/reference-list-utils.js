/**
 * Reference List Module (Task 5.2)
 * - Generate bibliography section
 * - Automate reference sorting
 * - Automate reference formatting
 * - Duplicate detection routines
 *
 * Reuses the citation database and APA formatting engine from
 * citation-utils.js (Task 5.1) so both modules stay in sync.
 */

import {
  getCitationDatabase,
  formatReferenceEntry,
  getDisambiguationLetter,
} from "./citation-utils";

const REFERENCES_HEADING = "References";

// ---------------------------------------------------------------------
// Task 5.2 - Automate reference sorting
// ---------------------------------------------------------------------

/**
 * Sorts citations alphabetically by first author's last name, then first
 * name, then year (standard APA reference list ordering).
 */
export function sortCitations(citations) {
  return [...citations].sort((a, b) => {
    const lastA = (a.authors[0]?.lastName || "").toLowerCase();
    const lastB = (b.authors[0]?.lastName || "").toLowerCase();
    if (lastA !== lastB) return lastA.localeCompare(lastB);

    const firstA = (a.authors[0]?.firstName || "").toLowerCase();
    const firstB = (b.authors[0]?.firstName || "").toLowerCase();
    if (firstA !== firstB) return firstA.localeCompare(firstB);

    return String(a.year).localeCompare(String(b.year));
  });
}

// ---------------------------------------------------------------------
// Task 5.2 - Duplicate detection routines
// ---------------------------------------------------------------------

/**
 * Finds groups of citations that likely refer to the same source: same
 * first author's last name, same year, and same (normalized) title.
 * Returns an array of arrays — each inner array is one duplicate group
 * with 2+ entries.
 */
export function findDuplicateCitations(citations) {
  const buckets = {};

  citations.forEach((c) => {
    const key = [
      (c.authors[0]?.lastName || "").trim().toLowerCase(),
      String(c.year).trim(),
      c.title.trim().toLowerCase().replace(/\s+/g, " "),
    ].join("|");
    buckets[key] = buckets[key] || [];
    buckets[key].push(c);
  });

  return Object.values(buckets).filter((group) => group.length > 1);
}

// ---------------------------------------------------------------------
// Task 5.2 - Automate reference formatting
// ---------------------------------------------------------------------

/**
 * Formats the full, sorted reference list as an array of strings, ready
 * to insert into the document. Applies disambiguation letters (a, b, c)
 * where author+year clash, matching the in-text citations from Task 5.1.
 */
export function buildFormattedReferenceList(citations) {
  const sorted = sortCitations(citations);
  return sorted.map((citation) => {
    const letter = getDisambiguationLetter(citation, citations);
    return formatReferenceEntry(citation, letter);
  });
}

// ---------------------------------------------------------------------
// Task 5.2 - Generate bibliography section
// ---------------------------------------------------------------------

/**
 * Finds an existing "References" heading paragraph in the document, or
 * returns null if none exists yet.
 */
async function findReferencesHeading(context) {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items/text, items/style");
  await context.sync();

  return (
    paragraphs.items.find(
      (p) => p.text.trim().toLowerCase() === REFERENCES_HEADING.toLowerCase()
    ) || null
  );
}

/**
 * Generates (or regenerates) the References section at the end of the
 * document, formatted per APA with a hanging indent. If a "References"
 * heading already exists, everything after it is replaced; otherwise a
 * new heading + list is appended at the end of the document.
 */
export async function generateBibliography(context) {
  const citations = await getCitationDatabase(context);

  const duplicates = findDuplicateCitations(citations);
  if (duplicates.length > 0) {
    throw new Error(
      `Found ${duplicates.length} likely duplicate source(s) in the citation database. Resolve duplicates before generating the bibliography.`
    );
  }

  const entries = buildFormattedReferenceList(citations);

  let headingParagraph = await findReferencesHeading(context);

  if (headingParagraph) {
    // Remove existing entries after the heading up to the next heading
    // (style "Heading 1"/"Heading 2") or end of document.
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/text, items/style");
    await context.sync();

    const headingIndex = paragraphs.items.findIndex(
      (p) => p.text.trim().toLowerCase() === REFERENCES_HEADING.toLowerCase()
    );

    for (let i = headingIndex + 1; i < paragraphs.items.length; i++) {
      const p = paragraphs.items[i];
      if (p.style && p.style.toLowerCase().includes("heading")) break;
      p.delete();
    }
    await context.sync();
  } else {
    // No References section yet — append one at the end of the document.
    const body = context.document.body;
    headingParagraph = body.insertParagraph(REFERENCES_HEADING, Word.InsertLocation.end);
    headingParagraph.style = "Heading 1";
    await context.sync();
  }

  let insertAfter = headingParagraph;
  entries.forEach((entryText) => {
    const p = insertAfter.insertParagraph(entryText, Word.InsertLocation.after);
    p.paragraphFormat.leftIndent = 36; // 0.5" hanging indent setup
    p.paragraphFormat.firstLineIndent = -36;
    p.font.size = 12;
    insertAfter = p;
  });

  await context.sync();
  return { count: entries.length };
}
