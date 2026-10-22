/**
 * Header and footer for the report (section 3 onward).
 *
 * Header: "Author's name" on the left, short FYP title on the right.
 * Footer: subject code (e.g. FYP4132) on the left, "Page X of Y" on the right.
 *
 * Sections 1 and 2 (the Front Cover) must have NO header or footer. Office.js
 * has no "Link to Previous" switch, so a header written into section 3 would
 * also appear on the covers. Instead, sections 1 and 2 turn on "Different
 * First Page": each cover is a single page, so that page shows its own
 * (empty) First Page header/footer, while section 3 onward shows the normal
 * (Primary) one. This only holds while each cover stays one page long.
 *
 * Page numbers continue from the covers (pages 1 and 2), so the Title Page,
 * the first page of section 3, is page 3.
 *
 * Layout: one line written as Word XML (OOXML), not a table, so there are no
 * dotted table gridlines. The line is widened past the body margins with
 * negative indents, so it runs from 2 cm off the left page edge to 2 cm off
 * the right page edge - centred on the page. The right-hand
 * text uses an absolute right-aligned tab (w:ptab), pinned to that right edge.
 *
 * Both functions clear and rebuild, so the buttons can be clicked again to
 * change the text.
 */

/* global Word */

const REPORT_SECTION_INDEX = 2; // section 3

const inchToPt = (inch) => inch * 72;
const cmToPt = (cm) => cm * 28.346;

// Section 3 margins (see margin.js).
const LEFT_MARGIN_PT = inchToPt(1.52);
const RIGHT_MARGIN_PT = inchToPt(1.2);

// Where the header/footer text sits, measured from the page edges. The
// same distance on both sides, so the line is centred on the page.
// Change this one number to move both sides in or out.
const FROM_PAGE_EDGE_CM = 2.0;
const FROM_LEFT_EDGE_PT = cmToPt(FROM_PAGE_EDGE_CM);
const FROM_RIGHT_EDGE_PT = cmToPt(FROM_PAGE_EDGE_CM);

// Negative indents pull the line out past the body margins. Word measures
// indents in twentieths of a point (twips).
const toTwips = (pt) => Math.round(pt * 20);
const LEFT_INDENT_PT = -(LEFT_MARGIN_PT - FROM_LEFT_EDGE_PT);
const RIGHT_INDENT_PT = -(RIGHT_MARGIN_PT - FROM_RIGHT_EDGE_PT);
const LEFT_INDENT_TWIPS = toTwips(LEFT_INDENT_PT);
const RIGHT_INDENT_TWIPS = toTwips(RIGHT_INDENT_PT);

// Height of the one 10 pt line, so it can be centred in the margin.
const LINE_HEIGHT_PT = 11.5;

/**
 * Vertical centring: puts the header line in the middle of the TOP margin
 * (between the page edge and where the text starts), and the footer line
 * in the middle of the BOTTOM margin. Applied to every report section
 * (section 3 onward), using each section's own margin.
 */
async function centreInMargins(context, sections, which) {
  const report = sections.items.slice(REPORT_SECTION_INDEX);
  report.forEach((section) => section.pageSetup.load("topMargin,bottomMargin"));
  await context.sync();
  report.forEach((section) => {
    const setup = section.pageSetup;
    if (which === "header") {
      setup.headerDistance = Math.max(0, (setup.topMargin - LINE_HEIGHT_PT) / 2);
    } else {
      setup.footerDistance = Math.max(0, (setup.bottomMargin - LINE_HEIGHT_PT) / 2);
    }
  });
  await context.sync();
}

const FONT_NAME = "Arial";
const FONT_SIZE_HALF_POINTS = 20; // 10 pt

async function loadSections(context) {
  const sections = context.document.sections;
  sections.load("items");
  await context.sync();
  if (sections.items.length <= REPORT_SECTION_INDEX) {
    throw new Error("Generate the Front Cover first. The header and footer start at section 3.");
  }
  return sections;
}

// Keeps the covers (sections 1 and 2) free of any header/footer.
function hideOnCovers(sections) {
  sections.items.forEach((section, i) => {
    if (i < REPORT_SECTION_INDEX) {
      section.pageSetup.differentFirstPageHeaderFooter = true;
      section.getHeader("FirstPage").clear();
      section.getFooter("FirstPage").clear();
    } else if (i === REPORT_SECTION_INDEX) {
      // Section 3's first page (the Title Page) shows the normal header/footer.
      section.pageSetup.differentFirstPageHeaderFooter = false;
    }
  });
}

function escapeXml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

const RUN_PROPS =
  `<w:rPr><w:rFonts w:ascii="${FONT_NAME}" w:hAnsi="${FONT_NAME}" w:cs="${FONT_NAME}"/>` +
  `<w:b w:val="0"/><w:color w:val="000000"/>` +
  `<w:sz w:val="${FONT_SIZE_HALF_POINTS}"/><w:szCs w:val="${FONT_SIZE_HALF_POINTS}"/></w:rPr>`;

const textRun = (text) =>
  `<w:r>${RUN_PROPS}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;

// A Word field (PAGE, NUMPAGES). The "1" is only the value shown until Word
// updates the field, which it does by itself in headers and footers.
const fieldRun = (code) =>
  `<w:fldSimple w:instr=" ${code} \\* MERGEFORMAT "><w:r>${RUN_PROPS}<w:t>1</w:t></w:r></w:fldSimple>`;

// Absolute tab, right-aligned at the right indent, so the right-hand text
// always ends at the same spot however long the left-hand text is.
const RIGHT_TAB = `<w:r>${RUN_PROPS}<w:ptab w:relativeTo="indent" w:alignment="right" w:leader="none"/></w:r>`;

function lineOoxml(leftRunsXml, rightRunsXml) {
  const paragraph =
    `<w:p><w:pPr>` +
    `<w:ind w:left="${LEFT_INDENT_TWIPS}" w:right="${RIGHT_INDENT_TWIPS}"/>` +
    `<w:spacing w:before="0" w:after="0"/><w:jc w:val="left"/>` +
    `</w:pPr>${leftRunsXml}${RIGHT_TAB}${rightRunsXml}</w:p>`;

  return (
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships></pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraph}</w:body></w:document>` +
    `</pkg:xmlData></pkg:part></pkg:package>`
  );
}

/**
 * Clears a header/footer body and writes the single left/right line into it.
 * Removes any extra empty paragraph Word leaves behind, so the line is the
 * only thing in the header/footer.
 */
async function writeLine(context, body, ooxml) {
  body.clear();
  body.insertOoxml(ooxml, "Replace");
  await context.sync();

  const paragraphs = body.paragraphs;
  paragraphs.load("items/text");
  await context.sync();
  // Keep the first (our line); delete trailing empty ones.
  for (let i = paragraphs.items.length - 1; i > 0; i--) {
    if (!paragraphs.items[i].text.trim()) paragraphs.items[i].delete();
  }

  // Set the indents directly as well. When Word inserts XML into an existing
  // paragraph it keeps that paragraph's own settings, so the w:ind in the XML
  // can be dropped - which left the text stuck at the body margins.
  const line = paragraphs.items[0];
  line.leftIndent = LEFT_INDENT_PT;
  line.rightIndent = RIGHT_INDENT_PT;
  line.firstLineIndent = 0;
  line.spaceBefore = 0;
  line.spaceAfter = 0;
  line.alignment = "Left";
  await context.sync();
}

export async function setReportHeader(context, authorName, shortTitle) {
  const sections = await loadSections(context);
  hideOnCovers(sections);
  await context.sync();

  const header = sections.items[REPORT_SECTION_INDEX].getHeader("Primary");
  await writeLine(context, header, lineOoxml(textRun(authorName), textRun(shortTitle)));
  await centreInMargins(context, sections, "header");
}

export async function setReportFooter(context, subjectCode) {
  const sections = await loadSections(context);
  hideOnCovers(sections);
  await context.sync();

  const footer = sections.items[REPORT_SECTION_INDEX].getFooter("Primary");
  const right = textRun("Page ") + fieldRun("PAGE") + textRun(" of ") + fieldRun("NUMPAGES");
  await writeLine(context, footer, lineOoxml(textRun(subjectCode), right));
  await centreInMargins(context, sections, "footer");
}
