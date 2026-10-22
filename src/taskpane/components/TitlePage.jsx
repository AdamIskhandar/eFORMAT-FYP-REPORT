import React, { useState, useEffect } from "react";
import { Button, Badge, Dropdown, Option, Input, Field } from "@fluentui/react-components";
import { LockClosed16Regular, DocumentAdd20Regular, CheckmarkCircle20Filled } from "@fluentui/react-icons";
import { isReportMarginSet, applyUniformMargin } from "./Section/utils/margin";
import { fillInLinesOoxml, setFillInText, fitFillInPage } from "./Section/utils/fill-in-fields";
import { StepBody, Hint, SubHeading, Group, Actions, Status, Notice, GuideTable, GuideSteps } from "./ui/StepUI";

/* global Word */

// Title Page layout (first page of section 3), Arial Bold 16, centred:
//   PROJECT TITLE        - upper part of the page
//   STUDENT FULL NAME    - middle
//   COURSE NAME          - lower part (e.g. BACHELOR OF INFORMATION TECHNOLOGY ...)
//   YEAR                 - last line, right under the course name
// Each line is a fill-in field: RED placeholder until the student types,
// then their text is BLACK (see fill-in-fields.js).
// Vertical position comes from spaceBefore (points), kept so a title and a
// course name that each wrap onto 3 lines still fit on one page.
const TITLE_PAGE_LINES = [
  { field: "Project Title", tag: "eformat-tp-title", placeholder: "PROJECT TITLE", size: 16, spaceBefore: 100, bookmark: "bmTpProjectTitle" },
  { field: "Student Full Name", tag: "eformat-tp-name", placeholder: "STUDENT FULL NAME", size: 16, spaceBefore: 160, bookmark: "bmTpFullName" },
  { field: "Course Name", tag: "eformat-tp-course", placeholder: "COURSE NAME", size: 16, spaceBefore: 190, bookmark: "bmTpCourse" },
  // Keep this bookmark name: the Declaration step looks for it to find the
  // end of the Title Page.
  { field: "Year", tag: "eformat-tp-year", placeholder: "YEAR", size: 16, spaceBefore: 20, bookmark: "bmTpMonthYear" },
];

// The Title Page's text area (report margin: 1" top, 1.2" bottom, 1.52"
// left, 1.2" right on A4), less a small safety gap, and the page-break line
// after the year, which must stay on this page too.
const TITLE_PAGE_AREA = {
  widthPt: 595.3 - 1.52 * 72 - 1.2 * 72,
  heightPt: 841.9 - 1 * 72 - 1.2 * 72 - 10,
  extraPt: 14,
};

// Examples shown in the panel's text box for each line.
const EXAMPLES = {
  "eformat-tp-title": "e.g. MPP INFORMATION APP",
  "eformat-tp-name": "e.g. ADAM ISKHANDAR BIN MUHASLEE",
  "eformat-tp-course": "e.g. BACHELOR OF INFORMATION TECHNOLOGY (HONS)",
  "eformat-tp-year": "e.g. 2026",
};

// Bookmarks that mean a Title Page already exists (new and older versions).
const EXISTING_MARKERS = ["bmTpProjectTitle", "bmTpUniversity"];

const GUIDE_ROWS = [
  { element: "Project Title", format: "Arial Bold 16 - upper part of the page" },
  { element: "Student Full Name", format: "Arial Bold 16 - middle" },
  { element: "Course Name", format: "Arial Bold 16 - lower part" },
  { element: "Year", format: "Arial Bold 16 - last line" },
];

export default function TitlePage() {
  const [marginStatus, setMarginStatus] = useState("");
  const [formatStatus, setFormatStatus] = useState("");
  const [marginIsSet, setMarginIsSet] = useState(false);
  const [fieldTag, setFieldTag] = useState("");
  const [fieldText, setFieldText] = useState("");
  const [fillStatus, setFillStatus] = useState("");

  // Reads section 3's real margins on load, so the Title Page button unlocks
  // if the margin is already in place (and stays locked if it was undone).
  useEffect(() => {
    Word.run(async (context) => {
      setMarginIsSet(await isReportMarginSet(context));
    }).catch(() => {});
  }, []);

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

  // ---- Title Page ----

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

  async function titlePageExists(context) {
    const found = EXISTING_MARKERS.map((name) => {
      const range = context.document.getBookmarkRangeOrNullObject(name);
      range.load("isNullObject");
      return range;
    });
    await context.sync();
    return found.some((range) => !range.isNullObject);
  }

  async function generateTitlePage() {
    if (!marginIsSet) {
      setFormatStatus("Set the margin first.");
      return;
    }
    try {
      await Word.run(async (context) => {
        const section = await getTitlePageSection(context);

        if (await titlePageExists(context)) {
          setFormatStatus("The Title Page is already there. Use Fill in the Title Page below to change its text.");
          return;
        }

        // Built at the START of section 3, top to bottom, then a page break
        // so anything after the Title Page starts on the next page.
        section.body.insertOoxml(fillInLinesOoxml(TITLE_PAGE_LINES, 840200), Word.InsertLocation.start);
        await context.sync();

        const paragraphs = section.body.paragraphs;
        paragraphs.load("items");
        await context.sync();

        TITLE_PAGE_LINES.forEach((line, i) => {
          paragraphs.items[i].getRange("Content").insertBookmark(line.bookmark);
        });
        paragraphs.items[TITLE_PAGE_LINES.length - 1].insertBreak(Word.BreakType.page, Word.InsertLocation.after);
        await context.sync();

        setFormatStatus(
          "Title Page created. Now fill in each red line using Fill in the Title Page below: choose a line, type your text, click Set Text."
        );
      });
    } catch (err) {
      console.error(err);
      setFormatStatus(err.message || "Could not create the Title Page.");
    }
  }

  async function handleFill() {
    const line = TITLE_PAGE_LINES.find((l) => l.tag === fieldTag);
    if (!line || !fieldText.trim()) {
      setFillStatus("Choose the line and type the text first.");
      return;
    }
    try {
      await Word.run(async (context) => {
        const ok = await setFillInText(context, line, fieldText.trim(), line.size);
        if (!ok) {
          setFillStatus("Line not found. Generate the Title Page first.");
          return;
        }
        // Keep the Title Page on one page: long text -> smaller gaps.
        const spacing = await fitFillInPage(context, TITLE_PAGE_LINES, TITLE_PAGE_AREA);
        let note = "";
        if (spacing && !spacing.fits) {
          note = " This text is too long to fit on one page even with smaller gaps - please shorten it.";
        } else if (spacing && spacing.shrunk) {
          note = " The spacing was made smaller so the Title Page stays on one page.";
        }
        setFillStatus(`Updated: ${line.field}.${note}`);
        setFieldText("");
      });
    } catch (err) {
      console.error(err);
      setFillStatus(err.message || "Could not update the Title Page.");
    }
  }
  const fillTone = /shorten it\.$/.test(fillStatus)
    ? "warning"
    : fillStatus.startsWith("Updated")
    ? "success"
    : /first\.$/.test(fillStatus)
    ? "warning"
    : "error";
  const chosen = TITLE_PAGE_LINES.find((l) => l.tag === fieldTag);

  // Message colours: the success texts are fixed; anything else is a
  // reminder (missing step) or an error from Word.
  const marginTone = /^(Margin has been set|The margin was set)/.test(marginStatus)
    ? "success"
    : "error";
  const formatTone = /^(Title Page created|The Title Page is already there)/.test(formatStatus)
    ? "success"
    : /first\.$/.test(formatStatus)
      ? "warning"
      : "error";
  const created = formatTone === "success";

  return (
    <StepBody>
      <SubHeading>
        Margin {marginIsSet && <Badge appearance="tint" color="success" size="small">Set</Badge>}
      </SubHeading>
      <Actions stretch>
        <Button appearance={marginIsSet ? "secondary" : "primary"} onClick={applyMostFrontCoverMargin}>
          {marginIsSet ? "Margin Set" : "Set Margin"}
        </Button>
      </Actions>
      <Status text={marginStatus} tone={marginTone} />

      <SubHeading>Title Page</SubHeading>
      <GuideSteps
        steps={[
          <>Click <b>Set Margin</b> above (only once).</>,
          <>Click <b>Generate Title Page</b> (only once).</>,
          <>In <b>Fill in the Title Page</b> below, choose a line: Project Title, Student Full Name, Course Name or Year.</>,
          <>Type your text in the box.</>,
          <>Click <b>Set Text</b>. The red words change to your text in black.</>,
          <>Do steps 3 - 5 again for every line until no red text is left.</>,
        ]}
        note="Fill in the Title Page using the boxes here - you don't need to type on the page in Word."
      />
      {!marginIsSet && (
        <Hint>
          <LockClosed16Regular style={{ verticalAlign: "text-bottom" }} /> Set the margin above
          first to unlock this part.
        </Hint>
      )}
      <Actions stretch>
        <Button
          appearance={created ? "secondary" : "primary"}
          icon={created ? <CheckmarkCircle20Filled /> : <DocumentAdd20Regular />}
          onClick={generateTitlePage}
          disabled={!marginIsSet}
        >
          {created ? "Title Page Created" : "Generate Title Page"}
        </Button>
      </Actions>
      <Status text={formatStatus} tone={formatTone} />

      <Group>
        <SubHeading>Fill in the Title Page</SubHeading>
        <Field label="1. Choose the line">
          <Dropdown
            placeholder="Select the line to fill"
            value={chosen ? chosen.field : ""}
            selectedOptions={fieldTag ? [fieldTag] : []}
            onOptionSelect={(_, data) => {
              setFieldTag(data.optionValue);
              setFillStatus("");
            }}
            disabled={!marginIsSet}
          >
            {TITLE_PAGE_LINES.map((l) => (
              <Option key={l.tag} value={l.tag}>
                {l.field}
              </Option>
            ))}
          </Dropdown>
        </Field>
        <Field label="2. Type the text">
          <Input
            placeholder={chosen ? EXAMPLES[chosen.tag] : "Type here"}
            value={fieldText}
            onChange={(_, data) => setFieldText(data.value)}
            onKeyDown={(e) => e.key === "Enter" && handleFill()}
            disabled={!marginIsSet}
          />
        </Field>
        <Actions>
          <Button appearance="primary" onClick={handleFill} disabled={!marginIsSet}>
            Set Text
          </Button>
        </Actions>
        <Status text={fillStatus} tone={fillTone} />
      </Group>

      <Notice tone="info">
        The Title Page always goes on the first page after the cover. It can&apos;t be added to the
        cover pages or to later pages.
      </Notice>
      <GuideTable title="Title Page format" rows={GUIDE_ROWS} />
    </StepBody>
  );
}
