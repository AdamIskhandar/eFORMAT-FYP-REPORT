import React, { useState } from "react";
import { makeStyles, Button, Text } from "@fluentui/react-components";
import { updateCaptionList, FIRST_CAPTION_CHAPTER } from "../Section/utils/caption-list-utils";

/* global Word */

const useStyles = makeStyles({
  ContentWrapper: { paddingLeft: "10px" },
  heading: { color: "white", fontSize: "20px" },
  Wrapper: { paddingLeft: "25px", display: "flex", flexDirection: "column", gap: "10px" },
  buttonRow: { display: "flex", gap: "10px", flexWrap: "wrap", marginTop: "6px" },
  status: { display: "block", marginTop: "6px", color: "#9fd89f" },
  warn: { display: "block", marginTop: "6px", color: "#e0a458" },
  error: { display: "block", marginTop: "6px", color: "#e07a7a" },
  note: { color: "#c7c7c7" },
});

// Explains WHY nothing was listed, using what the scan actually saw.
function explainEmpty(label, d) {
  const lines = [];
  if (d.headings === 0) {
    lines.push(
      'No numbered headings were found. Captions get their number from the heading above them, so the document needs Heading 1/2/3 paragraphs such as "CHAPTER 3: ..." or "3.1 ...".'
    );
  }
  if (d.noHeadingAbove > 0) {
    lines.push(`${d.noHeadingAbove} caption(s) sit above the first numbered heading.`);
  }
  if (d.beforeFirstChapter > 0) {
    lines.push(
      `${d.beforeFirstChapter} caption(s) are before Chapter ${FIRST_CAPTION_CHAPTER}, which are ignored.`
    );
  }
  if (d.notCentered > 0) {
    lines.push(
      `${d.notCentered} "${label} ..." paragraph(s) were skipped because they aren't centered.`
    );
  }
  if (lines.length === 0) lines.push(`No centered "${label} N: ..." captions were found.`);
  return lines.join(" ");
}

export default function CaptionLists() {
  const styles = useStyles();
  const [figureMsg, setFigureMsg] = useState({ text: "", tone: "status" });
  const [tableMsg, setTableMsg] = useState({ text: "", tone: "status" });

  async function update(type, label, setMsg) {
    try {
      await Word.run(async (context) => {
        const { count, diagnostics } = await updateCaptionList(context, type);
        setMsg(
          count > 0
            ? {
                text: `Renumbered and listed ${count} ${type}${count === 1 ? "" : "s"} (${diagnostics.headings} numbered heading${diagnostics.headings === 1 ? "" : "s"} recognized).`,
                tone: "status",
              }
            : { text: explainEmpty(label, diagnostics), tone: "warn" }
        );
      });
    } catch (err) {
      console.error(err);
      setMsg({ text: err.message || `Could not update the List of ${label}s.`, tone: "error" });
    }
  }

  return (
    <div className={styles.ContentWrapper}>
      <h2 className={styles.heading}>7) List of Figures &amp; Tables</h2>
      <div className={styles.Wrapper}>
        <Text className={styles.note}>
          Captions are numbered by the heading they sit under (Figure 3.1, Figure 3.3.2, ...),
          starting from Chapter {FIRST_CAPTION_CHAPTER}. Update renumbers every caption and rebuilds
          the list. Press Ctrl+A then F9 in Word afterwards to refresh the page numbers.
        </Text>

        <div className={styles.buttonRow}>
          <Button
            appearance="primary"
            shape="circular"
            onClick={() => update("figure", "Figure", setFigureMsg)}
          >
            Update List of Figures
          </Button>
          <Button
            appearance="primary"
            shape="circular"
            onClick={() => update("table", "Table", setTableMsg)}
          >
            Update List of Tables
          </Button>
        </div>

        {figureMsg.text && <Text className={styles[figureMsg.tone]}>{figureMsg.text}</Text>}
        {tableMsg.text && <Text className={styles[tableMsg.tone]}>{tableMsg.text}</Text>}
      </div>
    </div>
  );
}
