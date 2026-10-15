import React, { useState } from "react";
import { makeStyles, tokens, Button, Input, Text, Field } from "@fluentui/react-components";
import { StepBody, Hint, Actions, Status } from "../ui/StepUI";

/* global Word */

const useStyles = makeStyles({
  previewLabel: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
  },
  previewBox: {
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
    padding: "16px",
    maxHeight: "250px",
    overflowY: "auto",
    backgroundColor: tokens.colorNeutralBackground3,
  },
  previewHeading: {
    display: "block",
    textAlign: "center",
    fontWeight: "bold",
    letterSpacing: "1px",
    marginBottom: "14px",
  },
  previewTitle: {
    display: "block",
    textAlign: "center",
    fontWeight: "bold",
    marginBottom: "14px",
    textTransform: "uppercase",
  },
  previewParagraph: {
    display: "block",
    textAlign: "justify",
    fontSize: "12px",
    lineHeight: "1.2",
    color: tokens.colorNeutralForeground2,
  },
});

// Same constants as Acknowledgements.jsx, so both headings sit the same way.
const MM_TO_PT = 2.834645669;
const HEADING_TOP_SPACING_MM = 0;

// Standard sample academic abstract content (~220 words)
const SAMPLE_ABSTRACT_CONTENT =
  "This study investigates the structural and functional properties of advanced nanomaterials optimized for modern engineering applications. \n Over the past decade, material science has transitioned towards utilizing automated design frameworks to synthesize polymers with superior thermal stability. \n In this research, a comprehensive methodology combining empirical data collections and algorithmic modeling was deployed. \n The primary objectives focused on expanding current limitations surrounding baseline operational stress and structural fatigue under fluctuating environments. \n Initial testing matrices revealed an unprecedented 14% increment in overall tensile resilience when subjected to targeted localized pressure zones. Furthermore, cross-examinations of secondary variables verified that alignment variations do not disrupt structural integrity. These discoveries provide critical insights into improving scalable architectures for high-capacity manufacturing cycles. Ultimately, this framework establishes an innovative benchmark for structural optimization, enabling future architectural systems to seamlessly integrate robust parameters without inflating development overheads. This work bridges the gap between purely theoretical simulation frameworks and material constraints observed in modern industry settings.";

export default function Abstract() {
  const styles = useStyles();

  const [thesisTitle, setThesisTitle] = useState("");
  const [status, setStatus] = useState(null); // { type: "success" | "error", message: string }

  // Same structure as ensureSkeleton() in Acknowledgements.jsx:
  // reset vertical alignment, empty line, heading, (title), content, page break.
  async function ensureSkeleton(context) {
    const marker = context.document.getBookmarkRangeOrNullObject("bmAbstractHead");
    marker.load("isNullObject");
    await context.sync();
    if (!marker.isNullObject) return false;

    const sectionsBefore = context.document.sections;
    sectionsBefore.load("items");
    await context.sync();
    const abstractSection = sectionsBefore.items[sectionsBefore.items.length - 1];

    try {
      // Same as Acknowledgements: keep this section at Top vertical alignment.
      abstractSection.pageSetup.verticalAlignment = "Top";
      await context.sync();
    } catch (vAlignError) {
      console.warn(
        "Section vertical alignment is not supported in this Word version/host.",
        vAlignError
      );
    }

    // 0. Empty line above the heading. On the Acknowledgements page the
    //    heading also has one empty line above it (left over at the end of
    //    the document), which is what gives it that space at the top. Copy
    //    that line's size and spacing so ABSTRACT sits at exactly the same
    //    height; if it can't be found, use a plain 12pt Arial empty line.
    const spacer = context.document.body.insertParagraph("", Word.InsertLocation.end);
    spacer.font.name = "Arial";
    spacer.font.size = 12;
    spacer.font.bold = false;
    spacer.alignment = "Centered";
    spacer.spaceBefore = 0;
    spacer.spaceAfter = 0;

    const ackHead = context.document.getBookmarkRangeOrNullObject("bmAckHead");
    ackHead.load("isNullObject");
    await context.sync();
    if (!ackHead.isNullObject) {
      const lineAboveAck = ackHead.paragraphs.getFirst().getPreviousOrNullObject();
      lineAboveAck.load("isNullObject,text,lineSpacing,spaceBefore,spaceAfter,font/name,font/size");
      await context.sync();
      if (!lineAboveAck.isNullObject && lineAboveAck.text.trim() === "") {
        spacer.font.name = lineAboveAck.font.name;
        spacer.font.size = lineAboveAck.font.size;
        spacer.lineSpacing = lineAboveAck.lineSpacing;
        spacer.spaceBefore = lineAboveAck.spaceBefore;
        spacer.spaceAfter = lineAboveAck.spaceAfter;
      }
    }

    // 1. Heading - same settings as the ACKNOWLEDGEMENTS heading
    const heading = context.document.body.insertParagraph("ABSTRACT", Word.InsertLocation.end);
    heading.font.name = "Arial";
    heading.alignment = "Centered";
    heading.font.size = 12;
    heading.font.bold = true;
    heading.font.allCaps = true;
    heading.spaceBefore = HEADING_TOP_SPACING_MM * MM_TO_PT;
    heading.spaceAfter = 18;
    heading.lineSpacing = 12;
    heading.getRange().insertBookmark("bmAbstractHead");

    // 2. Thesis title (Abstract only - centred, bold, caps)
    const title = context.document.body.insertParagraph("", Word.InsertLocation.end);
    title.font.name = "Arial";
    title.font.size = 12;
    title.font.bold = true;
    title.font.allCaps = true;
    title.alignment = "Centered";
    title.spaceBefore = 0;
    title.spaceAfter = 18;
    title.lineSpacing = 12;
    title.getRange().insertBookmark("bmAbstractTitle");

    // 3. Content - same style as the Acknowledgements content, but single
    //    spaced (lineSpacing 12) because the Abstract is single-spaced.
    const content = context.document.body.insertParagraph(
      SAMPLE_ABSTRACT_CONTENT,
      Word.InsertLocation.end
    );
    content.font.name = "Arial";
    content.font.size = 12;
    content.font.bold = false;
    content.font.allCaps = false;
    content.alignment = "Justified";
    content.spaceBefore = 0;
    content.spaceAfter = 0;
    content.lineSpacing = 12;
    content.getRange().insertBookmark("bmAbstractContent");

    await context.sync();

    // 4. End the page, same as Acknowledgements, so the next page starts clean.
    context.document.body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
    await context.sync();

    return true;
  }

  // Same idea as fillField() in Acknowledgements.jsx: every click re-applies
  // the formatting, so the page always matches even if it was edited.
  async function fillField(context, bookmark, text) {
    const range = context.document.getBookmarkRangeOrNullObject(bookmark);
    range.load("isNullObject");
    await context.sync();
    if (range.isNullObject) return;

    const paragraph = range.paragraphs.getFirst();
    if (text && text.trim()) {
      paragraph.getRange("Content").insertText(text.trim(), Word.InsertLocation.replace);
    }
    paragraph.font.name = "Arial";
    paragraph.font.size = 12;

    if (bookmark === "bmAbstractHead" || bookmark === "bmAbstractTitle") {
      paragraph.font.bold = true;
      paragraph.font.allCaps = true;
      paragraph.alignment = "Centered";
      paragraph.spaceBefore = bookmark === "bmAbstractHead" ? HEADING_TOP_SPACING_MM * MM_TO_PT : 0;
      paragraph.spaceAfter = 18;
      paragraph.lineSpacing = 12;
    } else {
      // The content can be several paragraphs (each \n starts a new one).
      const paragraphs = range.paragraphs;
      paragraphs.load("items");
      await context.sync();
      paragraphs.items.forEach((p) => {
        p.font.name = "Arial";
        p.font.size = 12;
        p.font.bold = false;
        p.font.allCaps = false;
        p.alignment = "Justified";
        p.spaceBefore = 0;
        p.spaceAfter = 0;
        p.lineSpacing = 12;
      });
    }
    await context.sync();
  }

  async function handleSetFormat() {
    setStatus(null);
    if (!thesisTitle.trim()) {
      setStatus({ type: "error", message: "Please fill in your Thesis Title before generating." });
      return;
    }

    try {
      await Word.run(async (context) => {
        const created = await ensureSkeleton(context);

        await fillField(context, "bmAbstractHead", "ABSTRACT");
        await fillField(context, "bmAbstractTitle", thesisTitle);
        // Content keeps whatever the student wrote - only the format is reset.
        await fillField(context, "bmAbstractContent", "");

        setStatus({
          type: "success",
          message: created
            ? "Abstract page inserted successfully onto a separate page."
            : "Abstract page title updated with the changes you typed.",
        });
      });
    } catch (err) {
      console.error("Abstract formatting failed: ", err);
      const debugText = err?.debugInfo?.message ? ` (${err.debugInfo.message})` : "";
      setStatus({
        type: "error",
        message: `Failed to insert Abstract: ${err.message || "Unknown error"}${debugText}`,
      });
    }
  }

  return (
    <StepBody>
      <Hint>
        Adds the Abstract page after Acknowledgements: heading, your thesis title, then the
        abstract text single-spaced. Your own abstract text is kept when you click again.
      </Hint>

      <Field label="Thesis Title" required>
        <Input
          placeholder="e.g. STRUCTURAL CHARACTERISTICS OF ADVANCED POLYMERS"
          value={thesisTitle}
          onChange={(_, data) => setThesisTitle(data.value)}
        />
      </Field>

      <Text className={styles.previewLabel}>Preview</Text>
      <div className={styles.previewBox}>
        <Text className={styles.previewHeading}>ABSTRACT</Text>
        <Text className={styles.previewTitle}>
          {thesisTitle.trim() ? thesisTitle : "[YOUR THESIS TITLE WILL APPEAR HERE]"}
        </Text>
        <Text className={styles.previewParagraph}>{SAMPLE_ABSTRACT_CONTENT}</Text>
      </div>

      <Actions>
        <Button appearance="primary" onClick={handleSetFormat}>
          Set Format
        </Button>
      </Actions>
      {status && <Status text={status.message} tone={status.type === "success" ? "success" : "error"} />}
    </StepBody>
  );
}
