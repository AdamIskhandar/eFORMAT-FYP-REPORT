import React, { useState, useEffect } from "react";
import { useId, Dropdown, Option, Button, Input, Field, Badge } from "@fluentui/react-components";
import { LockClosed16Regular } from "@fluentui/react-icons";
import { isReportMarginSet, applyUniformMargin } from "./Section/utils/margin";
import { StepBody, Hint, SubHeading, Actions, Status, Notice } from "./ui/StepUI";

/* global Word */

// Title Page layout, matching the UPTM sample (first page of section 3):
//   UNIVERSITI POLY-TECH MALAYSIA      Arial Bold 18 - first line
//   Faculty of Computing & Multimedia  Arial Bold 15 - second line
//   PROJECT TITLE                      Arial Bold 16 - red, centre of page
//   STUDENT'S FULL NAME / STUDENT ID   Arial Bold 16 - red
//   COURSE CODE (e.g. FYP4025)         Arial Bold 16 - red, near the bottom
//   MONTH YEAR (e.g. NOVEMBER 2025)    Arial Bold 16 - last line
// Vertical position comes from spacingBefore (points). The total is kept
// about 100pt under one page at the section 3 margin, so a title that wraps
// onto a second line still fits on the page.
const RED = "#FF0000";
const BLACK = "#000000";

const FORMAT_FIELDS = [
  {
    id: "university",
    label: "Name of University",
    bookmark: "bmTpUniversity",
    order: 1,
    defaultText: "UNIVERSITI POLY-TECH MALAYSIA",
    fontSize: 18,
    allCaps: true,
    spacingBefore: 30,
    spacingAfter: 0,
    color: BLACK,
  },
  {
    id: "faculty",
    label: "Faculty",
    bookmark: "bmTpFaculty",
    order: 2,
    defaultText: "Faculty of Computing & Multimedia",
    fontSize: 15,
    allCaps: false,
    spacingBefore: 0,
    spacingAfter: 0,
    color: BLACK,
  },
  {
    id: "projectTitle",
    label: "Project Title",
    bookmark: "bmTpProjectTitle",
    order: 3,
    defaultText: "PROJECT TITLE",
    fontSize: 16,
    allCaps: true,
    spacingBefore: 180,
    spacingAfter: 0,
    color: RED,
  },
  {
    id: "fullName",
    label: "Student's Full Name",
    bookmark: "bmTpFullName",
    order: 4,
    defaultText: "STUDENT'S FULL NAME",
    fontSize: 16,
    allCaps: true,
    spacingBefore: 50,
    spacingAfter: 0,
    color: RED,
  },
  {
    id: "studentId",
    label: "Student ID",
    bookmark: "bmTpStudentId",
    order: 5,
    defaultText: "STUDENT ID",
    fontSize: 16,
    allCaps: true,
    spacingBefore: 0,
    spacingAfter: 0,
    color: RED,
  },
  {
    id: "courseCode",
    label: "Course Code (e.g. FYP4025)",
    bookmark: "bmTpCourseCode",
    order: 6,
    defaultText: "COURSE CODE",
    fontSize: 16,
    allCaps: true,
    spacingBefore: 190,
    spacingAfter: 0,
    color: RED,
  },
  {
    id: "monthYear",
    label: "Month & Year of Submission (e.g. NOVEMBER 2025)",
    bookmark: "bmTpMonthYear",
    order: 7,
    defaultText: "MONTH YEAR",
    fontSize: 16,
    allCaps: true,
    spacingBefore: 14,
    spacingAfter: 0,
    color: BLACK,
  },
];

export default function TitlePage() {
  const dropdownId = useId("dropdown");

  const [selectedFieldId, setSelectedFieldId] = useState(null);
  const [fieldText, setFieldText] = useState("");
  const [marginStatus, setMarginStatus] = useState("");
  const [formatStatus, setFormatStatus] = useState("");
  const [marginIsSet, setMarginIsSet] = useState(false);

  // Reads section 3's real margins on load, so the Format controls unlock if
  // the margin is already in place (and stay locked if it was undone).
  useEffect(() => {
    Word.run(async (context) => {
      setMarginIsSet(await isReportMarginSet(context));
    }).catch(() => {});
  }, []);

  function getFormatData(_, data) {
    const field = FORMAT_FIELDS.find((f) => f.label === data.optionValue);
    setSelectedFieldId(field ? field.id : null);
    setFormatStatus("");
  }

  // ---- Margin: single source of truth for the whole document, set once ----

  async function applyMostFrontCoverMargin() {
    try {
      await Word.run(async (context) => {
        // Check the document itself, not a saved flag - so this works again
        // after the margin is undone with Ctrl+Z.
        if (await isReportMarginSet(context)) {
          setMarginStatus("The margin was set.");
          setMarginIsSet(true);
          return;
        }
        await applyUniformMargin(context);
        setMarginStatus("Margin has been set for section 3 (Title Page onward).");
        setMarginIsSet(true);
      });
    } catch (err) {
      console.error(err);
      setMarginStatus(err.message || "Could not set the margin.");
    }
  }

  // ---- Format / auto-placement ----

  // The Title Page always lives on the FIRST page of section 3. Sections 1
  // and 2 are the Front Cover, and section 3's later pages are the report.
  async function getTitlePageSection(context) {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    if (sections.items.length < 3) {
      throw new Error(
        "Generate the Front Cover first. The Title Page can only go on the first page of section 3."
      );
    }
    return sections.items[2];
  }

  async function ensureSkeleton(context) {
    const section = await getTitlePageSection(context);

    const marker = context.document.getBookmarkRangeOrNullObject(FORMAT_FIELDS[0].bookmark);
    marker.load("isNullObject");
    await context.sync();

    if (!marker.isNullObject) {
      // Already built - make sure it's really in section 3 (an older version
      // of this add-in appended it to the end of the document instead).
      const where = marker.compareLocationWith(section.body.getRange("Whole"));
      await context.sync();
      const inside = ["Inside", "InsideStart", "InsideEnd", "Equal"].includes(where.value);
      if (!inside) {
        throw new Error(
          "The Title Page text is outside the first page of section 3. Delete it, then click Set Format again."
        );
      }
      return;
    }

    // Built at the START of section 3, top to bottom, then a page break so
    // anything after the Title Page starts on the next page.
    const ordered = [...FORMAT_FIELDS].sort((a, b) => a.order - b.order);
    let cursor = null;
    ordered.forEach((field, i) => {
      const text = field.defaultText || "";
      const paragraph =
        i === 0
          ? section.body.insertParagraph(text, Word.InsertLocation.start)
          : cursor.insertParagraph(text, Word.InsertLocation.after);
      paragraph.font.name = "Arial";
      paragraph.font.bold = true;
      paragraph.font.italic = false;
      paragraph.font.size = field.fontSize;
      paragraph.font.allCaps = field.allCaps;
      paragraph.font.color = field.color;
      paragraph.alignment = "Centered";
      paragraph.spaceBefore = field.spacingBefore;
      paragraph.spaceAfter = field.spacingAfter;
      // Bookmark the text only (not the paragraph mark), so replacing it
      // later can never merge this line into the next one.
      paragraph.getRange("Content").insertBookmark(field.bookmark);
      cursor = paragraph;
    });
    cursor.insertBreak(Word.BreakType.page, Word.InsertLocation.after);

    await context.sync();
  }

  async function applyMostFrontCoverFormat() {
    if (!marginIsSet) {
      setFormatStatus("Set the margin first.");
      return;
    }
    if (!selectedFieldId || !fieldText.trim()) {
      setFormatStatus("Type the text and choose a format first.");
      return;
    }
    const field = FORMAT_FIELDS.find((f) => f.id === selectedFieldId);

    try {
      await Word.run(async (context) => {
        await ensureSkeleton(context);

        const bookmarkRange = context.document.getBookmarkRangeOrNullObject(field.bookmark);
        bookmarkRange.load("isNullObject");
        await context.sync();
        if (bookmarkRange.isNullObject) {
          setFormatStatus("Could not find that field's placeholder. Try again.");
          return;
        }

        // Replace only the line's text (keeps its paragraph mark and position),
        // style the new text, then re-attach the bookmark to it so the same
        // field can be updated again later.
        const paragraph = bookmarkRange.paragraphs.getFirst();
        const newText = paragraph
          .getRange("Content")
          .insertText(fieldText.trim(), Word.InsertLocation.replace);
        newText.font.name = "Arial";
        newText.font.bold = true;
        newText.font.italic = false;
        newText.font.size = field.fontSize;
        newText.font.allCaps = field.allCaps;
        newText.font.color = field.color;
        paragraph.alignment = "Centered";
        paragraph.spaceBefore = field.spacingBefore;
        paragraph.spaceAfter = field.spacingAfter;
        newText.insertBookmark(field.bookmark);
        await context.sync();

        setFormatStatus(`Inserted into: ${field.label}`);
      });

      setFieldText("");
    } catch (err) {
      console.error(err);
      setFormatStatus(err.message || "Could not set the format.");
    }
  }

  // Message colours: the success texts are fixed; anything else is a
  // reminder (missing input) or an error from Word.
  const marginTone = /^(Margin has been set|The margin was set)/.test(marginStatus)
    ? "success"
    : "error";
  const formatTone = formatStatus.startsWith("Inserted into")
    ? "success"
    : /first\.$/.test(formatStatus)
      ? "warning"
      : "error";

  return (
    <StepBody>
      <SubHeading>
        Margin {marginIsSet && <Badge appearance="tint" color="success" size="small">Set</Badge>}
      </SubHeading>
      <Hint>Sets the report margin (1&quot; / 1.2&quot; / 1.52&quot; / 1.2&quot;) from the Title Page onward.</Hint>
      <Actions stretch>
        <Button appearance={marginIsSet ? "secondary" : "primary"} onClick={applyMostFrontCoverMargin}>
          {marginIsSet ? "Margin Set" : "Set Margin"}
        </Button>
      </Actions>
      <Status text={marginStatus} tone={marginTone} />

      <SubHeading>Title Page text</SubHeading>
      <Notice tone="info">
        The Title Page always goes on the first page of section 3 (right after the cover). It
        can&apos;t be added to the cover pages or to later pages.
      </Notice>
      {!marginIsSet && (
        <Hint>
          <LockClosed16Regular style={{ verticalAlign: "text-bottom" }} /> Set the margin above
          first to unlock this part.
        </Hint>
      )}

      <Field label="1. Choose the line">
        <Dropdown
          id={`${dropdownId}-format`}
          placeholder="Select the line to fill"
          value={selectedFieldId ? FORMAT_FIELDS.find((f) => f.id === selectedFieldId).label : ""}
          onOptionSelect={getFormatData}
          disabled={!marginIsSet}
        >
          {FORMAT_FIELDS.map((field) => (
            <Option key={field.id} value={field.label}>
              {field.label}
            </Option>
          ))}
        </Dropdown>
      </Field>

      <Field label="2. Type the text">
        <Input
          placeholder="e.g. MPP INFORMATION APP"
          value={fieldText}
          onChange={(_, data) => setFieldText(data.value)}
          disabled={!marginIsSet}
        />
      </Field>

      <Actions>
        <Button appearance="primary" onClick={applyMostFrontCoverFormat} disabled={!marginIsSet}>
          Set Format
        </Button>
      </Actions>
      <Status text={formatStatus} tone={formatTone} />
    </StepBody>
  );
}
