import React, { useState } from "react";
import { makeStyles, Button, Input, Text, Field } from "@fluentui/react-components";
import { TARGET_MARGIN } from "../Section/utils/margin";

/* global Word */

const useStyles = makeStyles({
  ContentWrapper: {
    paddingLeft: "10px",
  },
  heading: {
    color: "white",
    fontSize: "20px",
  },
  DeclarationWrapper: {
    paddingLeft: "25px",
    display: "flex",
    flexDirection: "column",
    gap: "10px",
  },
  button: {
    marginTop: "10px",
  },
  wrapButtonSize: {
    display: "flex",
    justifyContent: "flex-end",
    marginRight: "40px",
  },
  status: {
    display: "block",
    marginTop: "6px",
    color: "#9fd89f",
  },
  errorStatus: {
    display: "block",
    marginTop: "6px",
    color: "#ff8080",
  },
});

const cmToPt = (cm) => cm * 28.346;

// Right tab stop position for the Date/Name line: page width minus left and
// right margins, so the name lands right at the text area's right edge -
// matching the sample where the name sits flush right on the same line as
// the date.
const A4_WIDTH_PT = cmToPt(21);
const NAME_TAB_STOP_PT = A4_WIDTH_PT - TARGET_MARGIN.left - TARGET_MARGIN.right;

const DEFAULT_HEADING_TEXT = "DECLARATION";
const DEFAULT_CONTENT_TEXT =
  "I hereby declare that the work in this thesis is my own except for quotations and summaries which have been duly acknowledged.";

export default function Declaration() {
  const styles = useStyles();

  const [date, setDate] = useState("");
  const [name, setName] = useState("");
  const [studentId, setStudentId] = useState("");
  const [status, setStatus] = useState("");
  const [error, setError] = useState("");

  // Builds the whole declaration skeleton once: a page break, the heading
  // (pre-filled), the content paragraph (pre-filled), a shared date/name
  // line with a right tab stop, and the ID line. Returns true if it just
  // built the skeleton, false if it already existed.
  async function ensureSkeleton(context) {
    const marker = context.document.getBookmarkRangeOrNullObject("bmDeclHeading");
    await context.sync();
    if (!marker.isNullObject) return false;

    // Section break (not a plain page break) so Declaration can have its
    // own vertical alignment. verticalAlignment is a per-SECTION property
    // in Word — setting it on a shared section would also vertically
    // center the Title Page before it (and anything after it), so
    // Declaration needs to live in its own section to be centered on its
    // own. This does mean an extra sync round-trip to grab a reference to
    // the newly created section, which the previous single-page-break
    // version didn't need.
    // context.document.body.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.end);
    await context.sync();

    const sectionsBefore = context.document.sections;
    sectionsBefore.load("items");
    await context.sync();
    const declarationSection = sectionsBefore.items[sectionsBefore.items.length - 1];

    // This is the actual "center of the page, vertically" setting.
    // Its docs page is only published under Microsoft's preview API
    // set, same as tabStops earlier, so it's wrapped defensively —
    // if it's unsupported in this Word build/host, Declaration still
    // gets created, just without vertical centering.
    try {
      declarationSection.pageSetup.verticalAlignment = "Center";
      await context.sync();
    } catch (vAlignError) {
      console.warn(
        "Section vertical alignment is not supported in this Word version/host — Declaration will use normal top alignment.",
        vAlignError
      );
    }

    const heading = context.document.body.insertParagraph(
      DEFAULT_HEADING_TEXT,
      Word.InsertLocation.end
    );
    heading.font.name = "Arial";
    heading.font.bold = true;
    heading.font.size = 12;
    heading.font.allCaps = true;
    heading.alignment = "Centered";
    heading.spaceBefore = 226.8;
    heading.spaceAfter = 0;
    heading.getRange().insertBookmark("bmDeclHeading");

    const content = context.document.body.insertParagraph(
      DEFAULT_CONTENT_TEXT,
      Word.InsertLocation.end
    );
    content.font.name = "Arial";
    content.font.bold = false;
    content.font.size = 12;
    content.font.allCaps = false;
    content.alignment = "Left";
    content.spaceBefore = 28;
    content.spaceAfter = 0;
    content.getRange().insertBookmark("bmDeclContent");

    // Shared date/name line: paragraph starts as just a tab character, with
    // a right tab stop so whatever gets typed into "Name" lands flush right,
    // and a bookmark at each end so Date and Name can be filled independently.
    const dateNameLine = context.document.body.insertParagraph("\t", Word.InsertLocation.end);
    dateNameLine.font.name = "Arial";
    dateNameLine.font.size = 12;
    dateNameLine.alignment = "Left";
    dateNameLine.spaceBefore = 113.4;
    dateNameLine.spaceAfter = 0;

    // paragraphFormat.tabStops is also a newer/preview Word JS API and may
    // not be available depending on the Word build/version. Wrapped so a
    // failure here doesn't take down anything queued after it.
    try {
      dateNameLine.paragraphFormat.tabStops.clear();
      dateNameLine.paragraphFormat.tabStops.add(NAME_TAB_STOP_PT, Word.TabAlignment.right);
    } catch (tabStopError) {
      console.warn(
        "Tab stops are not supported in this Word version/host — date/name line will use default tab behavior.",
        tabStopError
      );
    }

    dateNameLine.getRange("Start").insertBookmark("bmDeclDate");
    dateNameLine.getRange("End").insertBookmark("bmDeclName");

    const idLine = context.document.body.insertParagraph("", Word.InsertLocation.end);
    idLine.font.name = "Arial";
    idLine.font.size = 12;
    idLine.font.bold = false;
    idLine.alignment = "Right";
    idLine.spaceBefore = 0;
    idLine.spaceAfter = 0;
    idLine.getRange().insertBookmark("bmDeclId");

    await context.sync();

    // Close off the Declaration section with another section break, and
    // reset the new section back to top alignment, so whatever gets
    // added after Declaration (Acknowledgements, Abstract, etc.) isn't
    // vertically centered too.

    context.document.body.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.end);
    await context.sync();

    const sectionsAfter = context.document.sections;
    sectionsAfter.load("items");
    await context.sync();
    const nextSection = sectionsAfter.items[sectionsAfter.items.length - 1];

    try {
      nextSection.pageSetup.verticalAlignment = "Top";
      await context.sync();
    } catch (vAlignError) {
      console.warn(
        "Could not reset vertical alignment for the section after Declaration.",
        vAlignError
      );
    }

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

        await fillField(context, "bmDeclDate", date, { bold: false, allCaps: false });
        await fillField(context, "bmDeclName", name, { bold: false, allCaps: true });
        await fillField(context, "bmDeclId", studentId, { bold: false, allCaps: false });

        setStatus(
          created
            ? "Declaration page generated: heading, sample content, and any details you typed."
            : "Declaration updated with the details you typed."
        );
      });

      setDate("");
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
    <div className={styles.ContentWrapper}>
      <h2 className={styles.heading}>2) Declaration</h2>
      <div className={styles.DeclarationWrapper}>
        <p>Show Example Of Declaration Page.</p>
        <Text style={{ color: "#c7c7c7" }}>
          The DECLARATION heading and sample sentence are generated automatically. Fill in your own
          date, name, and student ID below, then click Set Format.
        </Text>

        <Field label="Date">
          <Input
            placeholder="e.g. 29 May 2025"
            value={date}
            onChange={(_, data) => setDate(data.value)}
          />
        </Field>

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

        <div className={styles.wrapButtonSize}>
          <Button className={styles.button} shape="circular" onClick={handleSetFormat}>
            Set Format
          </Button>
        </div>
        {status && <Text className={styles.status}>{status}</Text>}
        {error && <Text className={styles.errorStatus}>Error: {error}</Text>}
      </div>
    </div>
  );
}
