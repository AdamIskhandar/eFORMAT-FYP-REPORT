/**
 * WHAT'S NEW - the update notes students see in the panel.
 *
 * HOW TO ANNOUNCE AN UPDATE (before you push to production):
 *   1. Add a NEW entry at the TOP of UPDATES.
 *   2. Give it a NEW version number (bigger than the one below it),
 *      e.g. "1.1.0" -> "1.2.0". The banner only shows when the version
 *      changes, so a new number = every student sees it once.
 *   3. Write the changes in short, simple sentences (what the student
 *      will notice, not code details).
 *   4. Push. After Netlify deploys, each student sees a "What's new" banner
 *      the next time they open eFormat. It goes away when they click Got it.
 *
 * Older entries can stay - they are shown under "Earlier updates".
 */

export const UPDATES = [
  {
    version: "1.1.1",
    date: "22 October 2026",
    changes: [
      "Long project titles: the Front Cover and Title Page now stay on one page - the spacing gets smaller by itself when your title or name is long.",
      "Table of Contents: page numbers are correct right after Generate (before, some could all show the same page until you clicked Update).",
      "Table of Contents: List of Figures and List of Tables now show the right page.",
      "Update List of Figures no longer risks making a second List of Figures.",
      "The 'Checking every page number' bar only appears when something really changed.",
      "The success message after Generate Front Cover is green again (it showed in red by mistake).",
      "Chapter names in the 'Add under' list are shown on one line.",
      "The header and footer are centred on the page.",
    ],
  },
  {
    version: "1.1.0",
    date: "22 October 2026",
    changes: [
      "Every step now has a How to use guide - click it to see what to do.",
      "Front Cover and Title Page: fill in your details using the boxes in the panel. The red text on the page turns black.",
      "Title Page has a new order: Project Title, Student Name, Course Name, Year.",
      "Declaration: the date is set to today. Click the date box to pick another day from the calendar.",
      "Acknowledgements: every paragraph after the first one starts with an indent.",
      "Abstract: the abstract is one paragraph with 1.5 line spacing.",
      "Chapters are now numbered CHAPTER 1, CHAPTER 2, ... (no more Roman numbers).",
      "New: Add a chapter - type a title and click Add Chapter.",
      "Headings stop at 3 digits (2.3.4). A subheading under 2.3.4 is lettered a., b., c. and is not listed in the Table of Contents.",
      "Add Subheading is faster: click Update Table of Contents once when you finish adding.",
      "Update Table of Contents now checks every page number for you.",
      "Text you type over \"Write your content here.\" comes out in the right format (Arial 11, 1.5 spacing, justify).",
    ],
  },
];

export const LATEST = UPDATES[0];
