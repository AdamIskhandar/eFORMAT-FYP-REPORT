import React, { useState } from "react";
import { Button, Input, Field } from "@fluentui/react-components";
import { StepBody, Actions, Status, GuideSteps } from "../ui/StepUI";

/* global Word */

// One absolute tab, right-aligned at the right margin. Text after it ends
// exactly at the right margin - the same edge as the right-aligned ID line.
const RIGHT_TAB_OOXML =
  '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
  '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  "</Relationships></pkg:xmlData></pkg:part>" +
  '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>' +
  '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>' +
  '<w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r>' +
  "</w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>";

const DEFAULT_HEADING_TEXT = "DECLARATION";
const DEFAULT_CONTENT_TEXT =
  "I hereby declare that the work in this thesis is my own except for quotations and summaries which have been duly acknowledged.";

const MONTHS = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December",
];

// Today as "YYYY-MM-DD" (local time), the value a date picker uses.
function todayValue() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

// "2026-10-22" -> "22 October 2026" (the format written on the page).
function formatDate(value) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value || "");
  if (!m) return "";
  return `${Number(m[3])} ${MONTHS[Number(m[2]) - 1]} ${m[1]}`;
}

export default function Declaration() {

  // Starts on today's date, so the page always gets a date.
  const [date, setDate] = useState(todayValue);
  const [name, setName] = useState("");
  const [studentId, setStudentId] = useState("");
  const [status, setStatus] = useState("");
  const [error, setError] = useState("");

  // Builds the whole declaration skeleton once, on the page right after the
  // Title Page: heading (pre-filled), content paragraph (pre-filled), a
  // shared date/name line, and the ID line. Returns true if it just built
  // the skeleton, false if it already existed.
  //
  // Vertical centring is a per-SECTION setting in Word, so the Declaration
  // gets its own section, set to "Center", with the section after it set
  // back to "Top":
  //   section 3 = Title Page
  //   section 4 = Declaration (vertically centred)
  //   section 5 = everything after (normal top alignment)
  // New sections copy section 3's page setup, so they keep the report
  // margin, A4, header and footer, and the page numbering continues.
  async function ensureSkeleton(context) {
    const marker = context.document.getBookmarkRangeOrNullObject("bmDeclHeading");
    marker.load("isNullObject");
    await context.sync();
    if (!marker.isNullObject) return false;

    const titlePageEnd = context.document.getBookmarkRangeOrNullObject("bmTpMonthYear");
    titlePageEnd.load("isNullObject");
    await context.sync();
    if (titlePageEnd.isNullObject) {
      throw new Error(
        "Set up the Title Page first (click Generate Title Page in step 2). The Declaration goes on the page after it."
      );
    }

    const lastTitleLine = titlePageEnd.paragraphs.getFirst();

    // The Title Page ends with a page break, in the paragraph right after
    // its "MONTH YEAR" line. Remove it - the section break below starts the
    // next page instead; keeping both would leave a blank page in between.
    const breakPara = lastTitleLine.getNextOrNullObject();
    breakPara.load("isNullObject,text");
    await context.sync();
    if (!breakPara.isNullObject && breakPara.text.replace(/[\f\s]/g, "") === "") {
      breakPara.delete();
    }

    // End the Title Page section here; the Declaration starts section 4.
    lastTitleLine.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.after);
    await context.sync();

    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    if (sections.items.length < 4) {
      throw new Error("Could not create the Declaration section. Try again.");
    }
    const declarationSection = sections.items[3];

    // Build top to bottom at the start of section 4.
    const heading = declarationSection.body.insertParagraph(
      DEFAULT_HEADING_TEXT,
      Word.InsertLocation.start
    );
    heading.font.name = "Arial";
    heading.font.bold = true;
    heading.font.italic = false;
    heading.font.size = 12;
    heading.font.color = "#000000";
    heading.font.allCaps = true;
    heading.alignment = "Centered";
    heading.spaceBefore = 0; // vertical centring positions the block
    heading.spaceAfter = 0;
    heading.getRange().insertBookmark("bmDeclHeading");

    const content = heading.insertParagraph(DEFAULT_CONTENT_TEXT, Word.InsertLocation.after);
    content.font.name = "Arial";
    content.font.bold = false;
    content.font.size = 12;
    content.font.color = "#000000";
    content.font.allCaps = false;
    content.alignment = "Left";
    content.spaceBefore = 28;
    content.spaceAfter = 0;
    content.getRange().insertBookmark("bmDeclContent");

    // Shared date/name line: paragraph starts as just a tab character, with
    // a right tab stop so whatever gets typed into "Name" lands flush right,
    // and a bookmark at each end so Date and Name can be filled independently.
    // Shared date/name line: date on the left, name pushed to the right
    // margin by an absolute right-aligned tab (w:ptab), so the name ends at
    // the same right edge as the right-aligned Student ID line below it.
    // (paragraphFormat.tabStops isn't available in this Word version, which
    // is why the name used to sit right after the date.)
    const dateNameLine = content.insertParagraph("", Word.InsertLocation.after);
    dateNameLine.font.name = "Arial";
    dateNameLine.font.bold = false;
    dateNameLine.font.size = 12;
    dateNameLine.font.color = "#000000";
    dateNameLine.alignment = "Left";
    dateNameLine.spaceBefore = 113.4;
    dateNameLine.spaceAfter = 0;

    const paragraphs = declarationSection.body.paragraphs;
    paragraphs.load("items");
    await context.sync();
    const countBefore = paragraphs.items.length;

    dateNameLine.getRange("Start").insertOoxml(RIGHT_TAB_OOXML, Word.InsertLocation.replace);
    await context.sync();

    // Inserting XML can leave an extra empty paragraph behind; remove it so
    // the Student ID line stays directly under the name.
    const check = declarationSection.body.paragraphs;
    check.load("items");
    await context.sync();
    if (check.items.length > countBefore) {
      const extra = dateNameLine.getNextOrNullObject();
      extra.load("isNullObject,text");
      await context.sync();
      if (!extra.isNullObject && extra.text.trim() === "") extra.delete();
    }

    // Date goes before the tab, name after it.
    dateNameLine.getRange("Start").insertBookmark("bmDeclDate");
    dateNameLine.getRange("End").insertBookmark("bmDeclName");

    const idLine = dateNameLine.insertParagraph("", Word.InsertLocation.after);
    idLine.font.name = "Arial";
    idLine.font.size = 12;
    idLine.font.bold = false;
    idLine.font.color = "#000000";
    idLine.alignment = "Right";
    idLine.spaceBefore = 0;
    idLine.spaceAfter = 0;
    idLine.getRange().insertBookmark("bmDeclId");

    // End the Declaration section, so whatever comes next starts on a new
    // page in section 5.
    idLine.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.after);
    await context.sync();

    const after = context.document.sections;
    after.load("items");
    await context.sync();

    // Section 4 (Declaration) centred; section 5 onward back to top.
    after.items[3].pageSetup.verticalAlignment = "Center";
    for (let i = 4; i < after.items.length; i++) {
      after.items[i].pageSetup.verticalAlignment = "Top";
    }
    await context.sync();
    return true;
  }

  async function fillField(context, bookmark, text, { bold, allCaps }) {
    if (!text.trim()) return;

    const bookmarkRange = context.document.getBookmarkRangeOrNullObject(bookmark);
    await context.sync();
    if (bookmarkRange.isNullObject) return;

    bookmarkRange.insertText(text.trim(), Word.InsertLocation.replace);
    await context.sync();

    const updatedRange = context.document.getBookmarkRangeOrNullObject(bookmark);
    updatedRange.font.name = "Arial";
    updatedRange.font.size = 12;
    updatedRange.font.bold = bold;
    updatedRange.font.allCaps = allCaps;
    await context.sync();
  }

  async function handleSetFormat() {
    setError("");
    try {
      await Word.run(async (context) => {
        const created = await ensureSkeleton(context);

        await fillField(context, "bmDeclDate", formatDate(date), { bold: false, allCaps: false });
        await fillField(context, "bmDeclName", name, { bold: false, allCaps: true });
        await fillField(context, "bmDeclId", studentId, { bold: false, allCaps: false });

        setStatus(
          created
            ? "Declaration page generated: heading, sample content, and any details you typed."
            : "Declaration updated with the details you typed."
        );
      });

      setDate(todayValue());
      setName("");
      setStudentId("");
    } catch (err) {
      // Previously this error was swallowed silently — Word.run had no
      // catch, so a thrown error (e.g. an unsupported API) just discarded
      // the whole batch with no feedback in the UI at all.
      console.error("Declaration Set Format failed:", err);
      setError(err.message || "Something went wrong generating the Declaration page.");
    }
  }

  return (
    <StepBody>
      <GuideSteps
        steps={[
          <>Check the <b>Date</b> - it is set to today. Click it to choose another day.</>,
          <>Type your <b>Name</b> and <b>Student ID</b>.</>,
          <>Click <b>Set Format</b>. The Declaration page is added after the Title Page.</>,
          <>Made a mistake? Fix the box and click <b>Set Format</b> again.</>,
        ]}
        note="Do step 2 (Title Page) first - the Declaration goes on the page after it."
      />

      <Field
        label="Date"
        hint={date ? `Shows on the page as: ${formatDate(date)}` : "Pick a date"}
      >
        <Input type="date" value={date} onChange={(_, data) => setDate(data.value)} />
      </Field>
      <Actions>
        <Button size="small" appearance="subtle" onClick={() => setDate(todayValue())}>
          Use today's date
        </Button>
      </Actions>
      <Field label="Name">
        <Input
          placeholder="e.g. Arena bt Che Kasim"
          value={name}
          onChange={(_, data) => setName(data.value)}
        />
      </Field>
      <Field label="Student ID">
        <Input
          placeholder="e.g. P10444"
          value={studentId}
          onChange={(_, data) => setStudentId(data.value)}
        />
      </Field>

      <Actions>
        <Button appearance="primary" onClick={handleSetFormat}>
          Set Format
        </Button>
      </Actions>
      <Status text={status} tone="success" />
      <Status text={error} tone="error" />
    </StepBody>
  );
}
