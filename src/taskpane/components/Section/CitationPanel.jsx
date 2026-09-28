import React, { useState } from "react";
import {
  Button,
  Dropdown,
  Option,
  Input,
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
  insertCitation,
  validateCitations,
  getCitationDatabase,
  deleteCitation,
} from "../Section/utils/citation-utils";
import {
  generateBibliography,
  findDuplicateCitations,
} from "../Section/utils/reference-list-utils";

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
  buttonRow: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    flexWrap: "wrap",
  },
  authorRow: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
  },
  issuesList: {
    margin: 0,
    paddingLeft: tokens.spacingHorizontalL,
    color: tokens.colorPaletteRedForeground1,
    fontSize: tokens.fontSizeBase200,
  },
  dbList: {
    margin: 0,
    paddingLeft: tokens.spacingHorizontalL,
    fontSize: tokens.fontSizeBase200,
  },
  ContentWrapper: {
    paddingLeft: "10px",
  },
  contentCoverCitation: {
    paddingLeft: "25px",
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    // padding: tokens.spacingHorizontalM,
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
 * UI section for Task 5.1 (Citation Processing) and Task 5.2
 * (Reference List). Slots in as sections "11)" and "12)".
 */
export default function CitationReferencePanel() {
  const styles = useStyles();

  const [sourceType, setSourceType] = useState("journal");
  const [citeStyle, setCiteStyle] = useState("parenthetical");
  const [firstName, setFirstName] = useState("");
  const [lastName, setLastName] = useState("");
  const [year, setYear] = useState("");
  const [title, setTitle] = useState("");
  const [source, setSource] = useState("");
  const [volume, setVolume] = useState("");
  const [issue, setIssue] = useState("");
  const [pages, setPages] = useState("");
  const [publisher, setPublisher] = useState("");
  const [url, setUrl] = useState("");
  const [doi, setDoi] = useState("");

  const [issues11, setIssues11] = useState([]);
  const [issues, setIssues] = useState([]);
  const [dbEntries, setDbEntries] = useState([]);
  const [status11, setStatus11] = useState(null);
  const [status, setStatus] = useState(null);

  const runInWord = async (fn) => {
    try {
      // setStatus({ intent: "info", text: "Working..." });
      await Word.run(fn);
    } catch (err) {
      console.error(err);
      setStatus({ intent: "error", text: err.message });
    }
  };

  const buildCitationData = () => ({
    type: sourceType,
    authors: [{ firstName, lastName }],
    year,
    title,
    source,
    volume,
    issue,
    pages,
    publisher,
    url,
    doi,
  });

  const clearForm = () => {
    setFirstName("");
    setLastName("");
    setYear("");
    setTitle("");
    setSource("");
    setVolume("");
    setIssue("");
    setPages("");
    setPublisher("");
    setUrl("");
    setDoi("");
  };

  const handleInsertCitation = () => {
    if (!lastName.trim() || !year.trim() || !title.trim()) {
      setStatus11({ intent: "warning", text: "Author last name, year, and title are required." });
      return;
    }
    runInWord(async (context) => {
      const result = await insertCitation(context, buildCitationData(), citeStyle);
      setStatus11({ intent: "success", text: `Inserted citation: ${result.text}` });
      clearForm();
    });
  };

  const handleValidate = () => {
    runInWord(async (context) => {
      const result = await validateCitations(context);
      setIssues11(result.issues);
      setStatus11({
        intent: result.valid ? "success" : "warning",
        text: result.valid
          ? "All citations are consistent."
          : `Found ${result.issues.length} issue(s).`,
      });
    });
  };

  const handleLoadDatabase = () => {
    runInWord(async (context) => {
      const citations = await getCitationDatabase(context);
      setDbEntries(citations);
      const duplicates = findDuplicateCitations(citations);
      setStatus11({
        intent: duplicates.length > 0 ? "warning" : "success",
        text:
          duplicates.length > 0
            ? `Loaded ${citations.length} source(s) — ${duplicates.length} duplicate group(s) found.`
            : `Loaded ${citations.length} source(s). No duplicates found.`,
      });
    });
  };

  const handleDeleteEntry = (id) => {
    runInWord(async (context) => {
      const remaining = await deleteCitation(context, id);
      setDbEntries(remaining);
      setStatus11({ intent: "success", text: "Removed source from the database." });
    });
  };

  const handleGenerateBibliography = () => {
    runInWord(async (context) => {
      const result = await generateBibliography(context);
      setStatus({
        intent: "success",
        text: `Generated References section with ${result.count} entr${result.count === 1 ? "y" : "ies"}.`,
      });
    });
  };

  return (
    <div className={styles.ContentWrapper}>
      <h3 className={styles.heading}>11) Citation Processing</h3>

      <div className={styles.contentCoverCitation}>
        <Field label="Source type">
          <Dropdown
            value={sourceType}
            selectedOptions={[sourceType]}
            onOptionSelect={(_, data) => setSourceType(data.optionValue)}
          >
            <Option value="journal">Journal article</Option>
            <Option value="book">Book</Option>
            <Option value="chapter">Book chapter</Option>
            <Option value="website">Website</Option>
          </Dropdown>
        </Field>

        <div className={styles.authorRow}>
          <Field label="Author first name">
            <Input value={firstName} onChange={(_, d) => setFirstName(d.value)} />
          </Field>
          <Field label="Author last name">
            <Input value={lastName} onChange={(_, d) => setLastName(d.value)} />
          </Field>
        </div>

        <Field label="Year">
          <Input value={year} onChange={(_, d) => setYear(d.value)} placeholder="e.g. 2024" />
        </Field>

        <Field label="Title">
          <Input value={title} onChange={(_, d) => setTitle(d.value)} />
        </Field>

        <Field label={sourceType === "book" ? "Publisher" : "Source (journal / site / book title)"}>
          <Input value={source} onChange={(_, d) => setSource(d.value)} />
        </Field>

        {sourceType === "journal" && (
          <div className={styles.authorRow}>
            <Field label="Volume">
              <Input value={volume} onChange={(_, d) => setVolume(d.value)} />
            </Field>
            <Field label="Issue">
              <Input value={issue} onChange={(_, d) => setIssue(d.value)} />
            </Field>
            <Field label="Pages">
              <Input value={pages} onChange={(_, d) => setPages(d.value)} />
            </Field>
          </div>
        )}

        {sourceType === "chapter" && (
          <div className={styles.authorRow}>
            <Field label="Pages">
              <Input value={pages} onChange={(_, d) => setPages(d.value)} />
            </Field>
            <Field label="Publisher">
              <Input value={publisher} onChange={(_, d) => setPublisher(d.value)} />
            </Field>
          </div>
        )}

        {(sourceType === "journal" || sourceType === "website") && (
          <Field label={sourceType === "journal" ? "DOI (optional)" : "URL"}>
            <Input
              value={sourceType === "journal" ? doi : url}
              onChange={(_, d) => (sourceType === "journal" ? setDoi(d.value) : setUrl(d.value))}
            />
          </Field>
        )}

        <Field label="In-text style">
          <Dropdown
            value={citeStyle}
            selectedOptions={[citeStyle]}
            onOptionSelect={(_, data) => setCiteStyle(data.optionValue)}
          >
            <Option value="parenthetical">Parenthetical — (Smith, 2024)</Option>
            <Option value="narrative">Narrative — Smith (2024)</Option>
          </Dropdown>
        </Field>

        <div className={styles.buttonRow}>
          <Button appearance="primary" onClick={handleInsertCitation}>
            Insert Citation
          </Button>
          <Button appearance="secondary" onClick={handleValidate}>
            Validate Citations
          </Button>
          <Button appearance="secondary" onClick={handleLoadDatabase}>
            View Source Database
          </Button>
        </div>

        {issues.length > 0 && (
          <ul className={styles.issuesList}>
            {issues.map((issue, i) => (
              <li key={i}>{issue}</li>
            ))}
          </ul>
        )}

        {dbEntries.length > 0 && (
          <ul className={styles.dbList}>
            {dbEntries.map((c) => (
              <li key={c.id}>
                {c.authors[0]?.lastName} ({c.year}) — {c.title}{" "}
                <Button
                  size="small"
                  appearance="transparent"
                  onClick={() => handleDeleteEntry(c.id)}
                >
                  Remove
                </Button>
              </li>
            ))}
          </ul>
        )}

        {status11 && (
          <MessageBar intent={status11.intent === "warning" ? "warning" : status11.intent}>
            <MessageBarBody>
              <MessageBarTitle>
                {status11.intent === "error"
                  ? "Error"
                  : status11.intent === "warning"
                    ? "Notice"
                    : "Status"}
              </MessageBarTitle>
              {status11.text}
            </MessageBarBody>
            <MessageBarActions
              containerAction={
                <Button
                  appearance="transparent"
                  icon={<Dismiss16Regular />}
                  aria-label="Dismiss"
                  onClick={() => setStatus11(null)}
                />
              }
            />
          </MessageBar>
        )}
      </div>

      <h3 className={styles.heading}>12) Reference List</h3>
      <div className={styles.contentCoverCitation}>
        <div className={styles.buttonRow}>
          <Button appearance="primary" onClick={handleGenerateBibliography}>
            Generate / Update Bibliography
          </Button>
        </div>

        {status && (
          <MessageBar intent={status.intent === "warning" ? "warning" : status.intent}>
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
