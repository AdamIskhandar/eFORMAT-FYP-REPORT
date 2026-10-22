import React, { useState } from "react";
import { makeStyles, tokens, Button, Text } from "@fluentui/react-components";
import { StepBody, Actions, Status, GuideSteps } from "../ui/StepUI";

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
    maxHeight: "220px",
    overflowY: "auto",
    backgroundColor: tokens.colorNeutralBackground3,
  },
  previewHeading: {
    display: "block",
    textAlign: "center",
    fontWeight: "bold",
    letterSpacing: "1px",
    marginBottom: "18px", // Matches Word spaceAfter spacing
  },
  previewParagraph: {
    display: "block",
    textAlign: "justify",
    fontSize: "12px",
    lineHeight: "1.5", // Simulates 1.5 spacing visually
    color: tokens.colorNeutralForeground2,
    textIndent: "24px", // first-line indent, like the page
  },
});

const MM_TO_PT = 2.834645669;
const HEADING_TOP_SPACING_MM = 0;

// First-line indent for the Acknowledgements paragraphs: 0.5 inch
// (1.27 cm), Word's standard paragraph indent. The FIRST paragraph under
// the heading has no indent; every paragraph after it is indented. Because
// it is a paragraph setting, pressing Enter in an indented paragraph starts
// a new paragraph that is indented too.
const FIRST_LINE_INDENT_PT = 36;

const SAMPLE_LIST = [
  "First and foremost, I would like to express my sincere gratitude to my supervisor, Prof. Dr. [Supervisor Name], for the continuous guidance, encouragement, and support throughout the course of this research.",
  "I would also like to thank the academic and administrative staff of the [Faculty/Institute Name] for their technical assistance and for providing the facilities needed to complete this study.",
  "My appreciation also goes to my fellow labmates and friends at [Research Group/Lab Name] for their helpful suggestions and moral support during this journey.",
  "I gratefully acknowledge the financial support provided through [Grant Name/Number] and [Scholarship Name], which made this research possible.",
  "Finally, I would like to express my deepest appreciation to my family for their unwavering support, patience, and encouragement throughout my studies.",
];

// Body paragraph format: Arial 12, justified, 1.5 spacing, first-line
// indent (none for the first paragraph).
function applyBodyFormat(paragraph, isFirst) {
  paragraph.font.name = "Arial";
  paragraph.font.size = 12;
  paragraph.font.bold = false;
  paragraph.font.allCaps = false;
  paragraph.alignment = "Justified";
  paragraph.spaceBefore = 0;
  paragraph.spaceAfter = 0;
  paragraph.lineSpacing = 18; // 1.5 lines for Arial 12pt
  paragraph.leftIndent = 0;
  paragraph.firstLineIndent = isFirst ? 0 : FIRST_LINE_INDENT_PT;
}

async function ensureSkeleton(context) {
  const marker = context.document.getBookmarkRangeOrNullObject("bmAckHead");
  await context.sync();
  if (!marker.isNullObject) return false;

  // FIX 1: Insert section break BEFORE starting Acknowledgements so it moves to a separate page
  // and completely breaks away from the Declaration page's vertical alignment centering rules.
  // context.document.body.insertBreak(Word.BreakType.nextPage, Word.InsertLocation.end);
  await context.sync();

  const sectionsBefore = context.document.sections;
  sectionsBefore.load("items");
  await context.sync();
  const AcknowledmentSection = sectionsBefore.items[sectionsBefore.items.length - 1];

  try {
    // FIX 2: Explicitly force this section layout back to Top vertical alignment
    AcknowledmentSection.pageSetup.verticalAlignment = "Top";
    await context.sync();
  } catch (vAlignError) {
    console.warn(
      "Section vertical alignment is not supported in this Word version/host.",
      vAlignError
    );
  }

  // 1. Insert Heading (30mm from top of page)
  const heading = context.document.body.insertParagraph(
    "ACKNOWLEDGEMENTS",
    Word.InsertLocation.end
  );
  heading.font.name = "Arial";
  heading.alignment = "Centered";
  heading.font.size = 12;
  heading.font.bold = true;
  heading.font.allCaps = true;
  heading.spaceBefore = HEADING_TOP_SPACING_MM * MM_TO_PT; // Exactly 30mm spacing from page top
  heading.spaceAfter = 18; // Gap separating heading and content body
  heading.lineSpacing = 12;
  heading.getRange().insertBookmark("bmAckHead");

  // 2. Insert Content - one Word paragraph per sample paragraph, each
  //    the first one flush left, the rest with a first-line indent.
  let first = null;
  let last = null;
  SAMPLE_LIST.forEach((text) => {
    const paragraph = context.document.body.insertParagraph(text, Word.InsertLocation.end);
    applyBodyFormat(paragraph, !first);
    if (!first) first = paragraph;
    last = paragraph;
  });
  first.getRange().expandTo(last.getRange()).insertBookmark("bmAckCont");

  await context.sync();

  // 3. Close off section with a clean structural break so next pages don't get messed up
  context.document.body.insertBreak(Word.BreakType.nextPage, Word.InsertLocation.end);
  await context.sync();

  const sectionsAfter = context.document.sections;
  sectionsAfter.load("items");
  await context.sync();
  const nextSection = sectionsAfter.items[sectionsAfter.items.length - 1];

  try {
    nextSection.pageSetup.verticalAlignment = "Top";
    await context.sync();
  } catch (vAlignError) {
    console.warn("Could not reset vertical alignment for the section after.", vAlignError);
  }

  return true;
}

function Acknowledgements() {
  const styles = useStyles();
  const [status, setStatus] = useState(null);

  // Page already there: re-apply the format to the heading and to every
  // paragraph under it on the Acknowledgements page - the student's own
  // text is kept, and paragraphs they added get the indent as well.
  async function reformatExisting(context) {
    const headRange = context.document.getBookmarkRangeOrNullObject("bmAckHead");
    headRange.load("isNullObject");
    await context.sync();
    if (headRange.isNullObject) return 0;

    const heading = headRange.paragraphs.getFirst();
    heading.font.name = "Arial";
    heading.alignment = "Centered";
    heading.font.size = 12;
    heading.font.bold = true;
    heading.font.allCaps = true;
    heading.spaceBefore = HEADING_TOP_SPACING_MM * MM_TO_PT;
    heading.spaceAfter = 18;
    heading.lineSpacing = 12;
    heading.firstLineIndent = 0;

    // The Acknowledgements page is its own Word section.
    const paragraphs = headRange.sections.getFirst().body.paragraphs;
    paragraphs.load("items/text");
    await context.sync();

    const headIndex = paragraphs.items.findIndex(
      (p) => p.text.trim().toUpperCase() === "ACKNOWLEDGEMENTS"
    );
    let count = 0;
    paragraphs.items.slice(headIndex + 1).forEach((p) => {
      if (!p.text.trim()) return; // leave empty lines alone
      applyBodyFormat(p, count === 0); // first paragraph: no indent
      count++;
    });
    await context.sync();
    return count;
  }

  async function handleSetFormat() {
    setStatus(null);
    try {
      await Word.run(async (context) => {
        const created = await ensureSkeleton(context);
        const count = created ? SAMPLE_LIST.length : await reformatExisting(context);

        setStatus({
          type: "success",
          message: created
            ? "Acknowledgements page inserted on a separate page. Paragraphs after the first one start with an indent."
            : `Acknowledgements format applied to ${count} paragraph${count === 1 ? "" : "s"} (with indent). Your text was kept.`,
        });
      });
    } catch (err) {
      const debugInfo = err?.debugInfo ? ` (${err.debugInfo.message})` : "";
      setStatus({
        type: "error",
        message: `Failed to insert page: ${err?.message || "Unknown error"}${debugInfo}`,
      });
    }
  }

  return (
    <StepBody>
      <GuideSteps
        steps={[
          <>Click <b>Set Format</b>. The Acknowledgements page is added with sample text.</>,
          <>In Word, replace the sample text and the [bracketed] parts with your own words.</>,
          <>Press Enter to start a new paragraph - it is indented for you.</>,
          <>When you finish writing, click <b>Set Format</b> again to tidy the format. Your text is kept.</>,
        ]}
      />

      <Text className={styles.previewLabel}>Preview</Text>
      <div className={styles.previewBox}>
        <Text className={styles.previewHeading}>ACKNOWLEDGEMENTS</Text>
        {SAMPLE_LIST.map((text, i) => (
          <Text
            key={text}
            className={styles.previewParagraph}
            style={i === 0 ? { textIndent: 0 } : undefined}
          >
            {text}
          </Text>
        ))}
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

export default Acknowledgements;
