import React, { useState } from "react";
import { makeStyles, tokens, Button, Dropdown, Option, Input, Field } from "@fluentui/react-components";
import {
  Add20Regular,
  NumberSymbol20Regular,
  CheckmarkStarburst20Regular,
  ArrowSync20Regular,
  ArrowRight20Regular,
} from "@fluentui/react-icons";
import {
  insertCaption,
  renumberCaptions,
  validateCaptions,
  bookmarkToNumber,
} from "../Section/utils/word-caption-utils";
import {
  listBookmarks,
  navigateToBookmark,
  updateReferencesAfterRenumber,
} from "../Section/utils/word-crossref-utils";

import { StepBody, Hint, SubHeading, Group, Actions, Status } from "../ui/StepUI";

const useStyles = makeStyles({
  issuesList: {
    margin: 0,
    paddingLeft: "18px",
    color: tokens.colorPaletteRedForeground1,
    fontSize: tokens.fontSizeBase200,
  },
  twoCol: {
    display: "grid",
    gridTemplateColumns: "minmax(0, 1fr) minmax(0, 1fr)",
    gap: "8px",
  },
});

/**
 * UI section for Task 4.1 (Caption Automation) and Task 4.2 (Cross
 * Referencing), styled with Fluent UI v9 to match the rest of the add-in.
 */
export default function CaptionCrossRefPanel() {
  const styles = useStyles();

  const [captionType, setCaptionType] = useState("figure");
  const [captionText, setCaptionText] = useState("");
  const [issues, setIssues] = useState([]);

  const [refType, setRefType] = useState("figure");
  const [availableRefs, setAvailableRefs] = useState([]);
  const [selectedRef, setSelectedRef] = useState("");

  const [status, setStatus] = useState(null); // { intent: "success" | "error" | "info", text }
  const [status9, setStatus9] = useState(null); // { intent: "success" | "error" | "info", text }

  const runInWord = async (fn) => {
    try {
      // setStatus({ intent: "info", text: "Working..." });
      await Word.run(fn);
    } catch (err) {
      console.error(err);
      setStatus({ intent: "error", text: err.message });
    }
  };

  const handleInsertCaption = () => {
    if (!captionText.trim()) {
      setStatus9({ intent: "warning", text: "Enter caption text first." });
      return;
    }
    runInWord(async (context) => {
      const result = await insertCaption(context, captionType, captionText);
      // Adding a caption mid-chapter shifts the later ones; keep their references right.
      await updateReferencesAfterRenumber(context, captionType, result.numberMap);
      setStatus9({ intent: "success", text: `Inserted: ${result.text}` });
      setCaptionText("");
    });
  };

  const handleRenumber = () => {
    runInWord(async (context) => {
      const numberMap = await renumberCaptions(context, captionType);
      await updateReferencesAfterRenumber(context, captionType, numberMap);
      setStatus9({
        intent: "success",
        text: `Renumbered ${Object.keys(numberMap).length} ${captionType} caption(s) and synced references.`,
      });
    });
  };

  const handleValidate = () => {
    runInWord(async (context) => {
      const result = await validateCaptions(context, captionType);
      setIssues(result.issues);
      setStatus9({
        intent: result.valid ? "success" : "warning",
        text: result.valid
          ? `All ${result.count} ${captionType} caption(s) are valid.`
          : `Found ${result.issues.length} issue(s).`,
      });
    });
  };

  const handleRefreshRefs = () => {
    runInWord(async (context) => {
      const names = await listBookmarks(context, refType);
      setAvailableRefs(names);
      setSelectedRef(names[0] || "");
      setStatus({
        intent: names.length > 0 ? "success" : "warning",
        text:
          names.length > 0
            ? `Found ${names.length} ${refType} caption(s) available to reference.`
            : `No ${refType} captions found yet — insert one first.`,
      });
    });
  };

  // "Figure_3_1" -> "3.1"
  const numberFromBookmark = (bookmarkName) => bookmarkToNumber(refType, bookmarkName);
  const refLabel = (bookmarkName) =>
    `${refType === "figure" ? "Figure" : "Table"} ${numberFromBookmark(bookmarkName) || ""}`;

  const handleGoTo = () => {
    if (!selectedRef) {
      setStatus({ intent: "warning", text: "Refresh and pick a caption first." });
      return;
    }
    runInWord(async (context) => {
      await navigateToBookmark(context, selectedRef);
      setStatus({ intent: "success", text: `Navigated to ${refLabel(selectedRef)}.` });
    });
  };

  const typeLabel = (t) => (t === "figure" ? "Figure" : "Table");

  return (
    <StepBody>
      <Group>
        <SubHeading>Insert a caption</SubHeading>
        <Hint>
          Put the cursor on the picture&apos;s line (or inside the table), then insert. Numbers follow
          the chapter: the first figure in Chapter 3 is Figure 3.1. A table caption goes above the
          table, a figure caption below it.
        </Hint>
        <Field label="Type">
            <Dropdown
              value={typeLabel(captionType)}
              selectedOptions={[captionType]}
              onOptionSelect={(_, data) => setCaptionType(data.optionValue)}
            >
              <Option value="figure">Figure</Option>
              <Option value="table">Table</Option>
            </Dropdown>
          </Field>
        <Field label="Caption text">
          <Input
            value={captionText}
            onChange={(_, data) => setCaptionText(data.value)}
            placeholder="e.g. System architecture"
          />
        </Field>
        <Actions stretch>
          <Button appearance="primary" icon={<Add20Regular />} onClick={handleInsertCaption}>
            Insert {typeLabel(captionType)} Caption
          </Button>
        </Actions>
        <div className={styles.twoCol}>
          <Button size="small" icon={<NumberSymbol20Regular />} onClick={handleRenumber}>
            Renumber
          </Button>
          <Button size="small" icon={<CheckmarkStarburst20Regular />} onClick={handleValidate}>
            Check
          </Button>
        </div>
        {issues.length > 0 && (
          <ul className={styles.issuesList}>
            {issues.map((issue, i) => (
              <li key={i}>{issue}</li>
            ))}
          </ul>
        )}
        {status9 && (
          <Status text={status9.text} tone={status9.intent} onDismiss={() => setStatus9(null)} />
        )}
      </Group>

      <Group>
        <SubHeading>Find a caption</SubHeading>
        <Hint>Jump straight to any figure or table caption in your report. Load the captions, pick one, then click Go To Caption.</Hint>
        <Field label="Type">
            <Dropdown
              value={typeLabel(refType)}
              selectedOptions={[refType]}
              onOptionSelect={(_, data) => {
                setRefType(data.optionValue);
                setAvailableRefs([]);
                setSelectedRef("");
              }}
            >
              <Option value="figure">Figure</Option>
              <Option value="table">Table</Option>
            </Dropdown>
          </Field>
        <Actions stretch>
          <Button icon={<ArrowSync20Regular />} onClick={handleRefreshRefs}>
            Load {typeLabel(refType)} Captions
          </Button>
        </Actions>

        {availableRefs.length > 0 && (
          <>
            <Field label="Caption">
              <Dropdown
                value={selectedRef ? refLabel(selectedRef) : ""}
                selectedOptions={selectedRef ? [selectedRef] : []}
                onOptionSelect={(_, data) => setSelectedRef(data.optionValue)}
              >
                {availableRefs.map((name) => (
                  <Option key={name} value={name}>
                    {refLabel(name)}
                  </Option>
                ))}
              </Dropdown>
            </Field>
            <Actions>
              <Button appearance="primary" icon={<ArrowRight20Regular />} onClick={handleGoTo}>
                Go To Caption
              </Button>
            </Actions>
          </>
        )}

        {status && <Status text={status.text} tone={status.intent} onDismiss={() => setStatus(null)} />}
      </Group>
    </StepBody>
  );
}
