import React, { useState } from "react";
import { Button } from "@fluentui/react-components";
import { ImageMultiple20Regular, Table20Regular } from "@fluentui/react-icons";
import { StepBody, Hint, Actions, Status } from "../ui/StepUI";
import {
  updateCaptionList,
  refreshCaptionListPageNumbers,
  FIRST_CAPTION_CHAPTER,
} from "../Section/utils/caption-list-utils";

/* global Word */

// Message colour for each kind of message this panel sets.
const TONE = { status: "success", warn: "warning", error: "error" };

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
  const [figureMsg, setFigureMsg] = useState({ text: "", tone: "status" });
  const [tableMsg, setTableMsg] = useState({ text: "", tone: "status" });

  // Word finishes laying out the pages in the background after the update,
  // so refresh the page numbers again a moment later (twice, for long reports).
  function refreshPageNumbersLater() {
    [1500, 4000].forEach((delay) =>
      setTimeout(() => {
        Word.run((context) => refreshCaptionListPageNumbers(context)).catch((err) =>
          console.warn("Delayed page-number refresh failed (press F9 in Word).", err)
        );
      }, delay)
    );
  }

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
      refreshPageNumbersLater();
    } catch (err) {
      console.error(err);
      setMsg({ text: err.message || `Could not update the List of ${label}s.`, tone: "error" });
    }
  }

  return (
    <StepBody>
      <Hint>
        Captions are numbered by chapter: the first figure in Chapter 3 is Figure 3.1, the next
        Figure 3.2 (tables count separately). Update renumbers every caption and rebuilds the list.
        Page numbers refresh by themselves a moment later.
      </Hint>

      <Actions stretch>
        <Button
          appearance="primary"
          icon={<ImageMultiple20Regular />}
          onClick={() => update("figure", "Figure", setFigureMsg)}
        >
          Update List of Figures
        </Button>
        <Status text={figureMsg.text} tone={TONE[figureMsg.tone]} />
        <Button
          appearance="primary"
          icon={<Table20Regular />}
          onClick={() => update("table", "Table", setTableMsg)}
        >
          Update List of Tables
        </Button>
        <Status text={tableMsg.text} tone={TONE[tableMsg.tone]} />
      </Actions>
    </StepBody>
  );
}
