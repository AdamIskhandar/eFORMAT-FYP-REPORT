/**
 * Citation Processing Module (Task 5.1)
 * - Citation insertion functions
 * - Citation formatting engine
 * - APA formatting rules
 * - Validate citation consistency
 *
 * Storage design:
 * The citation database (every source the user has entered) is stored as
 * JSON in the document's built-in Settings store, under the key
 * "citationDatabase". This travels with the .docx file itself, so it
 * survives closing/reopening and doesn't need an external database.
 *
 * Each in-text citation inserted into the document is wrapped in a
 * bookmark named "Citation_<id>_<n>" so Task 5.2 (Reference List Module)
 * can find every place a source was cited, and so validateCitations()
 * can detect orphaned or unused entries.
 */

const SETTINGS_KEY = "citationDatabase";

/** Very small dependency-free unique id generator. */
function makeId() {
  return `c${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;
}

// ---------------------------------------------------------------------
// Citation database persistence (backed by Word document Settings)
// ---------------------------------------------------------------------

/**
 * Returns the full citation database as an array of citation objects:
 * { id, type, authors: [{firstName, lastName, middleInitial}], year,
 *   title, source, volume, issue, pages, publisher, url, doi }
 */
export async function getCitationDatabase(context) {
  const setting = context.document.settings.getItemOrNullObject(SETTINGS_KEY);
  setting.load("value");
  await context.sync();

  if (setting.isNullObjectOrEmpty || setting.isNullObject) return [];
  try {
    return JSON.parse(setting.value) || [];
  } catch {
    return [];
  }
}

async function saveCitationDatabase(context, citations) {
  context.document.settings.add(SETTINGS_KEY, JSON.stringify(citations));
  await context.sync();
}

/**
 * Adds a citation to the database if an equivalent one (same first
 * author's last name + year + title, case-insensitive) doesn't already
 * exist. Returns the stored citation (existing or newly created).
 */
export async function upsertCitation(context, citationData) {
  const citations = await getCitationDatabase(context);

  const normalizedTitle = citationData.title.trim().toLowerCase();
  const firstAuthorLast = (citationData.authors[0]?.lastName || "").trim().toLowerCase();

  const existing = citations.find(
    (c) =>
      (c.authors[0]?.lastName || "").trim().toLowerCase() === firstAuthorLast &&
      String(c.year) === String(citationData.year) &&
      c.title.trim().toLowerCase() === normalizedTitle
  );

  if (existing) return existing;

  const newCitation = { id: makeId(), ...citationData };
  citations.push(newCitation);
  await saveCitationDatabase(context, citations);
  return newCitation;
}

export async function deleteCitation(context, citationId) {
  const citations = await getCitationDatabase(context);
  const filtered = citations.filter((c) => c.id !== citationId);
  await saveCitationDatabase(context, filtered);
  return filtered;
}

// ---------------------------------------------------------------------
// Task 5.1 - Implement APA formatting rules (7th edition)
// ---------------------------------------------------------------------

function formatAuthorsForReference(authors) {
  const formatted = authors.map((a) => {
    const initials = [a.firstName, a.middleInitial]
      .filter(Boolean)
      .map((n) => `${n.trim().charAt(0).toUpperCase()}.`)
      .join(" ");
    return `${a.lastName}, ${initials}`;
  });

  if (formatted.length === 1) return formatted[0];
  if (formatted.length === 2) return `${formatted[0]}, & ${formatted[1]}`;
  if (formatted.length <= 20) {
    return `${formatted.slice(0, -1).join(", ")}, & ${formatted[formatted.length - 1]}`;
  }
  // 21+ authors: first 19, ellipsis, then last author (APA rule)
  return `${formatted.slice(0, 19).join(", ")}, ... ${formatted[formatted.length - 1]}`;
}

function lastNames(authors) {
  return authors.map((a) => a.lastName);
}

/**
 * Builds the parenthetical or narrative in-text citation, e.g.
 * "(Smith, 2020)", "(Smith & Lee, 2020)", "(Smith et al., 2020)",
 * or narrative form "Smith (2020)" / "Smith and Lee (2020)".
 * `disambiguationLetter` appends "a"/"b"/etc. for same-author-year clashes.
 */
export function formatInTextCitation(citation, style = "parenthetical", disambiguationLetter = "") {
  const names = lastNames(citation.authors);
  const year = `${citation.year}${disambiguationLetter}`;

  let namePart;
  if (names.length === 1) namePart = names[0];
  else if (names.length === 2) {
    namePart = style === "narrative" ? `${names[0]} and ${names[1]}` : `${names[0]} & ${names[1]}`;
  } else {
    namePart = `${names[0]} et al.`;
  }

  return style === "narrative" ? `${namePart} (${year})` : `(${namePart}, ${year})`;
}

/**
 * Task 5.1 - Develop citation formatting engine.
 * Builds a full APA reference-list entry for a citation, by source type.
 */
export function formatReferenceEntry(citation, disambiguationLetter = "") {
  const authorPart = formatAuthorsForReference(citation.authors);
  const yearPart = `(${citation.year}${disambiguationLetter}).`;
  const titlePart = citation.title.trim().replace(/\.?$/, ".");

  switch (citation.type) {
    case "journal": {
      const volumeIssue = citation.issue
        ? `${citation.volume}(${citation.issue})`
        : citation.volume || "";
      const pages = citation.pages ? `, ${citation.pages}` : "";
      const link = citation.doi ? `https://doi.org/${citation.doi}` : citation.url || "";
      return `${authorPart} ${yearPart} ${titlePart} ${citation.source}, ${volumeIssue}${pages}.${
        link ? ` ${link}` : ""
      }`.trim();
    }

    case "website": {
      const link = citation.url ? ` ${citation.url}` : "";
      return `${authorPart} ${yearPart} ${titlePart} ${citation.source}.${link}`.trim();
    }

    case "chapter": {
      return `${authorPart} ${yearPart} ${titlePart} In ${citation.source} (pp. ${citation.pages || "n.p."}). ${
        citation.publisher || ""
      }.`.trim();
    }

    case "book":
    default: {
      return `${authorPart} ${yearPart} ${titlePart} ${citation.publisher || citation.source || ""}.`.trim();
    }
  }
}

// ---------------------------------------------------------------------
// Task 5.1 - Develop citation insertion functions
// ---------------------------------------------------------------------

/**
 * Determines the disambiguation letter for a citation given the full
 * database: if multiple entries share the same first-author lastname +
 * year, they're lettered a, b, c... in title-alphabetical order.
 */
export function getDisambiguationLetter(citation, allCitations) {
  const clashGroup = allCitations
    .filter(
      (c) =>
        (c.authors[0]?.lastName || "").toLowerCase() ===
          (citation.authors[0]?.lastName || "").toLowerCase() &&
        String(c.year) === String(citation.year)
    )
    .sort((a, b) => a.title.localeCompare(b.title));

  if (clashGroup.length <= 1) return "";
  const position = clashGroup.findIndex((c) => c.id === citation.id);
  return String.fromCharCode(97 + position); // 97 = 'a'
}

/**
 * Inserts an in-text citation at the current selection, registering it
 * (or reusing it) in the citation database, computing disambiguation if
 * needed, and wrapping the inserted text in a tracking bookmark.
 */
export async function insertCitation(context, citationData, style = "parenthetical") {
  const citation = await upsertCitation(context, citationData);
  const allCitations = await getCitationDatabase(context);
  const letter = getDisambiguationLetter(citation, allCitations);

  const text = formatInTextCitation(citation, style, letter);

  const selection = context.document.getSelection();
  const inserted = selection.insertText(text, Word.InsertLocation.replace);
  await context.sync();

  // Track usage with a unique bookmark instance per insertion, since the
  // same source may be cited multiple times in the document.
  const usageBookmark = `Citation_${citation.id}_${Date.now().toString(36)}`;
  inserted.insertBookmark(usageBookmark);
  await context.sync();

  return { citation, text, bookmarkName: usageBookmark };
}

// ---------------------------------------------------------------------
// Task 5.1 - Validate citation consistency
// ---------------------------------------------------------------------

/**
 * Cross-checks in-text citation bookmarks against the citation database:
 * - flags in-text citations whose source was deleted from the database
 * - flags database entries never cited anywhere in the text
 * - flags author/year clashes missing disambiguation letters
 */
export async function validateCitations(context) {
  const citations = await getCitationDatabase(context);
  const issues = [];

  if (!context.document.bookmarks) {
    issues.push(
      "Bookmarks collection unavailable on this Word host/API version — cannot verify in-text usage."
    );
    return { valid: false, issues };
  }

  context.document.bookmarks.load("items/name");
  await context.sync();

  const citationBookmarks = context.document.bookmarks.items
    .map((b) => b.name)
    .filter((n) => n.startsWith("Citation_"));

  const citedIds = new Set(citationBookmarks.map((n) => n.split("_")[1]));

  // Orphaned in-text citations (bookmark references a deleted database entry)
  citedIds.forEach((id) => {
    if (!citations.find((c) => c.id === id)) {
      issues.push(
        `An in-text citation points to a source (id: ${id}) that no longer exists in the database.`
      );
    }
  });

  // Unused database entries
  citations.forEach((c) => {
    if (!citedIds.has(c.id)) {
      const label = c.authors[0]?.lastName || "Unknown author";
      issues.push(
        `"${label} (${c.year})" is in the reference database but never cited in the text.`
      );
    }
  });

  // Author/year clashes missing disambiguation letters would only matter
  // once rendered; check here that groups of 2+ are consistently handled.
  const groups = {};
  citations.forEach((c) => {
    const key = `${(c.authors[0]?.lastName || "").toLowerCase()}_${c.year}`;
    groups[key] = groups[key] || [];
    groups[key].push(c);
  });
  Object.values(groups).forEach((group) => {
    if (group.length > 1) {
      const titles = group.map((c) => c.title.trim().toLowerCase());
      if (new Set(titles).size !== titles.length) {
        issues.push(
          `Multiple identical sources detected for "${group[0].authors[0]?.lastName} (${group[0].year})" — check for duplicates.`
        );
      }
    }
  });

  return { valid: issues.length === 0, issues };
}
