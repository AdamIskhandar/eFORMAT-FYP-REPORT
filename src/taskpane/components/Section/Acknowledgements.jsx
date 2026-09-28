import React, { useState } from "react";
import { makeStyles, Button, Text } from "@fluentui/react-components";

/* global Word */

const useStyles = makeStyles({
  ContentWrapper: {
    paddingLeft: "10px",
  },
  heading: {
    color: "white",
    fontSize: "20px",
  },
  AcknowledgementsWrapper: {
    paddingLeft: "25px",
  },
  description: {
    marginBottom: "10px",
  },
  previewBox: {
    border: "1px solid #4a4a4a",
    borderRadius: "4px",
    padding: "16px", // Balanced padding
    marginBottom: "14px",
    maxHeight: "220px",
    overflowY: "auto",
    backgroundColor: "#2b2b2b",
  },
  previewHeading: {
    display: "block",
    textAlign: "center",
    fontWeight: "bold",
    letterSpacing: "1px",
    marginBottom: "18px", // Matches Word spaceAfter spacing
    color: "#ffffff",
  },
  previewParagraph: {
    display: "block",
    textAlign: "justify",
    fontSize: "12px",
    lineHeight: "1.5", // Simulates 1.5 spacing visually
    color: "#d6d6d6",
  },
  wrapButtonSize: {
    display: "flex",
    justifyContent: "flex-end",
    marginRight: "40px",
    maxWidth: "250px",
    minWidth: "250px",
  },
  button: {
    marginTop: "10px",
  },
  statusText: {
    display: "block",
    marginTop: "8px",
    fontSize: "12px",
  },
  statusSuccess: {
    color: "#3fbf60",
  },
  statusError: {
    color: "#e04b4b",
  },
});

const MM_TO_PT = 2.834645669;
const HEADING_TOP_SPACING_MM = 0;

const SAMPLE_PARAGRAPHS =
  "First and foremost, I would like to express my sincere gratitude to my supervisor, Prof. Dr. [Supervisor Name], for the continuous guidance, encouragement, and support throughout the course of this research. \n I would also like to thank the academic and administrative staff of the [Faculty/Institute Name] for their technical assistance and for providing the facilities needed to complete this study. \n My appreciation also goes to my fellow labmates and friends at [Research Group/Lab Name] for their helpful suggestions and moral support during this journey. \n I gratefully acknowledge the financial support provided through [Grant Name/Number] and [Scholarship Name], which made this research possible. \n Finally, I would like to express my deepest appreciation to my family for their unwavering support, patience, and encouragement throughout my studies.";

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

  // 2. Insert Content
  const content = context.document.body.insertParagraph(SAMPLE_PARAGRAPHS, Word.InsertLocation.end);
  content.font.name = "Arial";
  content.font.size = 12;
  content.font.bold = false;
  content.font.allCaps = false;
  content.alignment = "Justified";
  content.spaceBefore = 0;
  content.spaceAfter = 0;
  content.lineSpacing = 18; // Double-spacing standard (1.5 line equivalent for Arial 12pt)
  content.getRange().insertBookmark("bmAckCont");

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

  async function fillField(context, bookmark, text, { bold, allCaps }) {
    if (!text.trim()) return;

    const bookmarkRange = context.document.getBookmarkRangeOrNullObject(bookmark);
    await context.sync();
    if (bookmarkRange.isNullObject) return;

    bookmarkRange.insertText(text.trim(), Word.InsertLocation.replace);
    await context.sync();

    if (bookmark === "bmAckHead") {
      const updatedRange = context.document.getBookmarkRangeOrNullObject(bookmark);
      updatedRange.font.name = "Arial";
      updatedRange.alignment = "Centered";
      updatedRange.font.size = 12;
      updatedRange.font.bold = true;
      updatedRange.font.allCaps = true;
      updatedRange.spaceBefore = HEADING_TOP_SPACING_MM * MM_TO_PT;
      updatedRange.spaceAfter = 18;
      updatedRange.lineSpacing = 12;
      await context.sync();
    } else {
      const currentParagraph = context.document.getBookmarkRangeOrNullObject(bookmark);
      currentParagraph.font.name = "Arial";
      currentParagraph.font.size = 12;
      currentParagraph.font.bold = false;
      currentParagraph.font.allCaps = false;
      currentParagraph.alignment = "Justified";
      currentParagraph.spaceBefore = 0;
      currentParagraph.spaceAfter = 0;
      currentParagraph.lineSpacing = 18;
      await context.sync();
    }
  }

  async function handleSetFormat() {
    setStatus(null);
    try {
      await Word.run(async (context) => {
        const created = await ensureSkeleton(context);

        await fillField(context, "bmAckHead", "ACKNOWLEDGEMENTS", { bold: true, allCaps: true });
        await fillField(context, "bmAckCont", SAMPLE_PARAGRAPHS, { bold: false, allCaps: false });

        setStatus({
          type: "success",
          message: created
            ? "Acknowledgements page inserted on a separate page."
            : "Acknowledgements page updated successfully.",
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
    <div className={styles.ContentWrapper}>
      <h2 className={styles.heading}>3) Acknowledgements</h2>
      <div className={styles.AcknowledgementsWrapper}>
        <p className={styles.description}>
          Click "Set Format" to auto-generate the Acknowledgements page at your cursor position.
        </p>

        <div className={styles.previewBox}>
          <Text className={styles.previewHeading}>ACKNOWLEDGEMENTS</Text>
          <Text className={styles.previewParagraph}>{SAMPLE_PARAGRAPHS}</Text>
        </div>

        <div className={styles.wrapButtonSize}>
          <Button
            className={styles.button}
            shape="circular"
            appearance="primary"
            onClick={handleSetFormat}
          >
            Set Format
          </Button>
        </div>

        {status && (
          <Text
            className={`${styles.statusText} ${
              status.type === "success" ? styles.statusSuccess : styles.statusError
            }`}
          >
            {status.message}
          </Text>
        )}
      </div>
    </div>
  );
}

export default Acknowledgements;
