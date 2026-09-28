import React, { useState } from "react";
import {
  Button,
  Dropdown,
  Option,
  Input,
  Label,
  Field,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  MessageBarActions,
  Divider,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { Dismiss16Regular } from "@fluentui/react-icons";
import {
  insertCaption,
  renumberCaptions,
  validateCaptions,
} from "../Section/utils/word-caption-utils";
import {
  insertReference,
  listBookmarks,
  navigateToBookmark,
  updateReferencesAfterRenumber,
} from "../Section/utils/word-crossref-utils";

const useStyles = makeStyles({
  panel: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingHorizontalM,
  },
  sectionTitle: {
    marginTop: tokens.spacingVerticalL,
    marginBottom: 0,
  },
  row: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
  },
  buttonRow: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    flexWrap: "wrap",
  },
  buttonRow2: {
    display: "flex",
    gap: tokens.spacingVerticalL,
    flexWrap: "wrap",
    marginTop: "10px",
  },
  issuesList: {
    margin: 0,
    paddingLeft: tokens.spacingHorizontalL,
    color: tokens.colorPaletteRedForeground1,
    fontSize: tokens.fontSizeBase200,
    marginTop: "10px",
  },
  ContentWrapper: {
    paddingLeft: "10px",
  },
  contentCover: {
    paddingLeft: "25px",
  },
  heading: {
    color: "white",
    fontSize: "20px",
  },
  label: {
    marginBottom: "10px",
  },
  wrapButtonSize: {
    display: "flex",
    justifyContent: "flex-end",
    marginRight: "40px",
    maxWidth: "250px",
    minWidth: "250px",
  },
  messageBar: {
    marginTop: "10px",
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

  const numberFromBookmark = (bookmarkName) => {
    const match = bookmarkName && bookmarkName.match(/_(\d+)$/);
    return match ? parseInt(match[1], 10) : null;
  };

  const handleInsertReference = () => {
    const number = numberFromBookmark(selectedRef);
    if (!number) {
      setStatus({ intent: "warning", text: "Refresh and pick a caption to reference first." });
      return;
    }
    runInWord(async (context) => {
      await insertReference(context, refType, number, "see ");
      setStatus({ intent: "success", text: `Inserted reference to ${refType} ${number}.` });
    });
  };

  const handleGoTo = () => {
    if (!selectedRef) {
      setStatus({ intent: "warning", text: "Refresh and pick a caption first." });
      return;
    }
    runInWord(async (context) => {
      await navigateToBookmark(context, selectedRef);
      setStatus({ intent: "success", text: `Navigated to ${selectedRef}.` });
    });
  };

  return (
    <div className={styles.ContentWrapper}>
      <h3 className={styles.heading}>9) Caption Automation</h3>

      <div className={styles.contentCover}>
        <div className={styles.label}>
          <Field label="Type">
            <Dropdown
              value={captionType === "figure" ? "Figure" : "Table"}
              selectedOptions={[captionType]}
              onOptionSelect={(_, data) => setCaptionType(data.optionValue)}
            >
              <Option value="figure">Figure</Option>
              <Option value="table">Table</Option>
            </Dropdown>
          </Field>
        </div>

        <div className={styles.label}>
          <Field label="Caption text">
            <Input
              value={captionText}
              onChange={(_, data) => setCaptionText(data.value)}
              placeholder="e.g. System architecture diagram"
            />
          </Field>
        </div>

        <div className={styles.buttonRow}>
          <Button appearance="primary" onClick={handleInsertCaption}>
            Insert Caption
          </Button>
          <Button appearance="secondary" onClick={handleRenumber}>
            Renumber All
          </Button>
          <Button appearance="secondary" onClick={handleValidate}>
            Validate Captions
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
          <MessageBar
            intent={status9.intent === "warning" ? "warning" : status9.intent}
            className={styles.messageBar}
          >
            <MessageBarBody>
              <MessageBarTitle>
                {status9.intent === "error"
                  ? "Error"
                  : status9.intent === "warning"
                    ? "Notice"
                    : "Status"}
              </MessageBarTitle>
              {status9.text}
            </MessageBarBody>
            <MessageBarActions
              containerAction={
                <Button
                  appearance="transparent"
                  icon={<Dismiss16Regular />}
                  aria-label="Dismiss"
                  onClick={() => setStatus9(null)}
                />
              }
            />
          </MessageBar>
        )}
      </div>

      <h3 className={styles.heading}>10) Cross Referencing</h3>

      <div className={styles.contentCover}>
        <div className={styles.label}>
          <Field label="Reference type">
            <Dropdown
              value={refType === "figure" ? "Figure" : "Table"}
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
        </div>

        <div className={styles.buttonRow2}>
          <Button appearance="secondary" onClick={handleRefreshRefs}>
            Refresh Available Captions
          </Button>
        </div>
        {availableRefs.length > 0 && (
          <>
            <div className={styles.label}>
              <Field label="Select caption">
                <Dropdown
                  value={selectedRef.replace("_", " ")}
                  selectedOptions={selectedRef ? [selectedRef] : []}
                  onOptionSelect={(_, data) => setSelectedRef(data.optionValue)}
                >
                  {availableRefs.map((name) => (
                    <Option key={name} value={name}>
                      {name.replace("_", " ")}
                    </Option>
                  ))}
                </Dropdown>
              </Field>
            </div>

            <div className={styles.buttonRow2}>
              <Button appearance="primary" onClick={handleInsertReference}>
                Insert Reference
              </Button>
              <Button appearance="secondary" onClick={handleGoTo}>
                Go To
              </Button>
            </div>
          </>
        )}

        {status && (
          <MessageBar
            intent={status.intent === "warning" ? "warning" : status.intent}
            className={styles.messageBar}
          >
            <MessageBarBody>
              <MessageBarTitle>
                {status.intent === "error"
                  ? "Error"
                  : status.intent === "warning"
                    ? "Notice"
                    : "Status"}
              </MessageBarTitle>
              {status.text}
            </MessageBarBody>
            <MessageBarActions
              containerAction={
                <Button
                  appearance="transparent"
                  icon={<Dismiss16Regular />}
                  aria-label="Dismiss"
                  onClick={() => setStatus(null)}
                />
              }
            />
          </MessageBar>
        )}
      </div>
    </div>
  );
}
