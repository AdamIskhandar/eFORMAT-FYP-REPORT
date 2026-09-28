import React, { useState } from "react";
import { Button, Input, Field, makeStyles, tokens } from "@fluentui/react-components";
import {
  TextNumberListLtr24Regular,
  Notepad24Regular,
  Quote24Regular,
  TextQuote24Regular,
} from "@fluentui/react-icons";

/* global Word */

const useStyles = makeStyles({
  section: { display: "flex", flexDirection: "column", gap: "12px" },
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
  buttonRow: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    flexWrap: "wrap",
  },
  buttonReferences: {
    marginTop: "10px",
  },
});

export default function ReferenceModules() {
  const styles = useStyles();
  const [footnoteText, setFootnoteText] = useState("");
  const [citation, setCitation] = useState("");

  async function insertTableOfContents() {
    await Word.run(async (context) => {
      const range = context.document.body.insertParagraph("", Word.InsertLocation.start).getRange();
      range.insertField(
        Word.InsertLocation.replace,
        Word.FieldType.toc,
        '\\o "1-3" \\h \\z \\u',
        false
      );
      await context.sync();
    });
  }

  async function insertFootnote() {
    if (!footnoteText.trim()) return;
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.insertFootnote(footnoteText);
      await context.sync();
    });
    setFootnoteText("");
  }

  // Simplified citation insert. A full bibliography/citation manager would
  // need a separate data store (source list) keyed by citation id.
  async function insertCitation() {
    if (!citation.trim()) return;
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.insertText(`(${citation})`, Word.InsertLocation.replace);
      await context.sync();
    });
    setCitation("");
  }

  return (
    <div className={styles.ContentWrapper}>
      <h2 className={styles.heading}>15) References</h2>

      <div className={styles.contentCover}>
        <Button icon={<TextNumberListLtr24Regular />} onClick={insertTableOfContents}>
          Insert table of contents
        </Button>

        <div className={styles.label}>
          <Field label="Footnote text">
            <Input
              value={footnoteText}
              onChange={(_, data) => setFootnoteText(data.value)}
              placeholder="e.g. Source: internal report, 2026"
            />
          </Field>
          <div className={styles.buttonReferences}>
            <Button icon={<Notepad24Regular />} onClick={insertFootnote}>
              Insert footnote at cursor
            </Button>
          </div>
        </div>

        <Field label="Citation (author, year)">
          <Input
            value={citation}
            onChange={(_, data) => setCitation(data.value)}
            placeholder="e.g. Smith, 2025"
          />
        </Field>
        <div className={styles.buttonReferences}>
          <Button icon={<TextQuote24Regular />} onClick={insertCitation}>
            Insert inline citation
          </Button>
        </div>
      </div>
    </div>
  );
}
