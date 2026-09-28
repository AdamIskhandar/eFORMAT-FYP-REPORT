import React, { useState } from "react";
import { makeStyles, Button, Input, Text, Field } from "@fluentui/react-components";

/* global Word */

const useStyles = makeStyles({
  ContentWrapper: {
    paddingLeft: "10px",
  },
  heading: {
    color: "white",
    fontSize: "20px",
  },
  AbstractWrapper: {
    paddingLeft: "25px",
    display: "flex",
    flexDirection: "column",
    gap: "10px",
  },
  description: {
    color: "#c7c7c7",
    marginBottom: "5px",
  },
  previewBox: {
    border: "1px solid #4a4a4a",
    borderRadius: "4px",
    padding: "16px",
    marginBottom: "14px",
    maxHeight: "250px",
    overflowY: "auto",
    backgroundColor: "#2b2b2b",
  },
  previewHeading: {
    display: "block",
    textAlign: "center",
    fontWeight: "bold",
    letterSpacing: "1px",
    marginBottom: "14px",
    color: "#ffffff",
  },
  previewTitle: {
    display: "block",
    textAlign: "center",
    fontWeight: "bold",
    marginBottom: "14px",
    color: "#ffffff",
    textTransform: "uppercase",
  },
  previewParagraph: {
    display: "block",
    textAlign: "justify",
    fontSize: "12px",
    lineHeight: "1.2",
    color: "#d6d6d6",
  },
  wrapButtonSize: {
    display: "flex",
    justifyContent: "flex-end",
    marginRight: "40px",
  },
  button: {
    marginTop: "10px",
  },
  statusSuccess: {
    display: "block",
    marginTop: "6px",
    color: "#9fd89f",
  },
  statusError: {
    display: "block",
    marginTop: "6px",
    color: "#ff8080",
  },
});

// 1 mm = 2.834645669 points
const MM_TO_PT = 2.834645669;
const HEADING_TOP_SPACING_MM = 30;

// Standard sample academic abstract content (~220 words)
const SAMPLE_ABSTRACT_CONTENT =
  "This study investigates the structural and functional properties of advanced nanomaterials optimized for modern engineering applications. \n Over the past decade, material science has transitioned towards utilizing automated design frameworks to synthesize polymers with superior thermal stability. \n In this research, a comprehensive methodology combining empirical data collections and algorithmic modeling was deployed. \n The primary objectives focused on expanding current limitations surrounding baseline operational stress and structural fatigue under fluctuating environments. \n Initial testing matrices revealed an unprecedented 14% increment in overall tensile resilience when subjected to targeted localized pressure zones. Furthermore, cross-examinations of secondary variables verified that alignment variations do not disrupt structural integrity. These discoveries provide critical insights into improving scalable architectures for high-capacity manufacturing cycles. Ultimately, this framework establishes an innovative benchmark for structural optimization, enabling future architectural systems to seamlessly integrate robust parameters without inflating development overheads. This work bridges the gap between purely theoretical simulation frameworks and material constraints observed in modern industry settings.";

export default function Abstract() {
  const styles = useStyles();

  const [thesisTitle, setThesisTitle] = useState("");
  const [status, setStatus] = useState(null); // { type: "success" | "error", message: string }

  async function ensureSkeleton(context) {
    // Check if the Abstract page has already been generated
    const marker = context.document.getBookmarkRangeOrNullObject("bmAbstractHead");
    await context.sync();
    if (!marker.isNullObject) return false;

    // Create a new page section break after Acknowledgements
    // context.document.body.insertBreak(Word.BreakType.nextPage, Word.InsertLocation.end);
    await context.sync();

    // Reset this section layout to normal top alignment
    const currentSections = context.document.sections;
    currentSections.load("items");
    await context.sync();
    const abstractSection = currentSections.items[currentSections.items.length - 1];

    try {
      abstractSection.pageSetup.verticalAlignment = "Top";
      await context.sync();
    } catch (e) {
      console.warn("Vertical alignment configuration unsupported by this Word host.", e);
    }

    // 1. Insert ABSTRACT Heading (30mm from top of page)
    const heading = context.document.body.insertParagraph("ABSTRACT", Word.InsertLocation.end);
    heading.font.name = "Arial";
    heading.font.size = 12;
    heading.font.bold = true;
    heading.font.allCaps = true;
    heading.alignment = "Centered";
    heading.spaceBefore = HEADING_TOP_SPACING_MM * MM_TO_PT; // 30 mm spacing
    heading.spaceAfter = 18; // Gap before the title block
    heading.getRange().insertBookmark("bmAbstractHead");

    // 2. Insert Thesis Title Block (Centered, bold, uppercase)
    const titleParagraph = context.document.body.insertParagraph("", Word.InsertLocation.end);
    titleParagraph.font.name = "Arial";
    titleParagraph.font.size = 12;
    titleParagraph.font.bold = true;
    titleParagraph.font.allCaps = true;
    titleParagraph.alignment = "Centered";
    titleParagraph.spaceBefore = 0;
    titleParagraph.spaceAfter = 18; // Gap before content paragraph
    titleParagraph.getRange().insertBookmark("bmAbstractTitle");

    // 3. Insert Abstract Content Paragraph (Justified, Single-spaced)
    const contentParagraph = context.document.body.insertParagraph(
      SAMPLE_ABSTRACT_CONTENT,
      Word.InsertLocation.end
    );
    contentParagraph.font.name = "Arial";
    contentParagraph.font.size = 12;
    contentParagraph.font.bold = false;
    contentParagraph.font.allCaps = false;
    contentParagraph.alignment = "Justified";
    contentParagraph.spaceBefore = 0;
    contentParagraph.spaceAfter = 0;
    contentParagraph.lineSpacing = 12; // Single spacing (12 points font size = 12 point line space)
    contentParagraph.getRange().insertBookmark("bmAbstractContent");

    await context.sync();
    return true;
  }

  async function updateField(context, bookmark, text) {
    if (!text.trim()) return;
    const range = context.document.getBookmarkRangeOrNullObject(bookmark);
    await context.sync();
    if (range.isNullObject) return;

    range.insertText(text.trim().toUpperCase(), Word.InsertLocation.replace);
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

        // Populate or update the title bookmark
        await updateField(context, "bmAbstractTitle", thesisTitle);

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
    <div className={styles.ContentWrapper}>
      <h2 className={styles.heading}>4) Abstract</h2>
      <div className={styles.AbstractWrapper}>
        <Text className={styles.description}>
          Provide your thesis title to safely auto-generate a single-spaced, beautifully formatted
          Abstract page directly after Acknowledgements.
        </Text>

        {/* Input field for custom title */}
        <Field label="Thesis Title" required>
          <Input
            placeholder="e.g. STRUCTURAL CHARACTERISTICS OF ADVANCED POLYMERS"
            value={thesisTitle}
            onChange={(_, data) => setThesisTitle(data.value)}
          />
        </Field>

        {/* Student Live UI Layout Preview */}
        <p style={{ margin: "5px 0 0 0", fontWeight: "600" }}>Document Preview Window:</p>
        <div className={styles.previewBox}>
          <Text className={styles.previewHeading}>ABSTRACT</Text>
          <Text className={styles.previewTitle}>
            {thesisTitle.trim() ? thesisTitle : "[YOUR THESIS TITLE WILL APPEAR HERE]"}
          </Text>
          <Text className={styles.previewParagraph}>{SAMPLE_ABSTRACT_CONTENT}</Text>
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
          <Text className={status.type === "success" ? styles.statusSuccess : styles.statusError}>
            {status.type === "success" ? "✔ " : "❌ "} {status.message}
          </Text>
        )}
      </div>
    </div>
  );
}
