import React from "react";
import { Button, makeStyles, tokens } from "@fluentui/react-components";
import {
  DocumentHeaderFooter24Regular,
  DocumentPageNumber24Regular,
  TableSimple24Regular,
  DocumentSplitHint24Regular,
} from "@fluentui/react-icons";

/* global Word */

const useStyles = makeStyles({
  section: { display: "flex", flexDirection: "column", gap: "10px" },
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
});

export default function DocumentModules() {
  const styles = useStyles();

  async function insertHeader() {
    await Word.run(async (context) => {
      const sections = context.document.sections;
      const header = sections.getFirst().getHeader(Word.HeaderFooterType.primary);
      header.insertText("Document Header", Word.InsertLocation.replace);
      await context.sync();
    });
  }

  async function insertPageNumber() {
    await Word.run(async (context) => {
      const sections = context.document.sections;
      const footer = sections.getFirst().getFooter(Word.HeaderFooterType.primary);
      footer.insertField(Word.InsertLocation.end, "PAGE");
      await context.sync();
    });
  }

  async function insertSectionBreak() {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.after);
      await context.sync();
    });
  }

  async function insertTable() {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      const table = range.insertTable(3, 3, Word.InsertLocation.after, [
        ["Col 1", "Col 2", "Col 3"],
        ["", "", ""],
        ["", "", ""],
      ]);
      table.styleBuiltIn = Word.BuiltInStyleName.gridTable1Light;
      await context.sync();
    });
  }

  return (
    <div className={styles.ContentWrapper}>
      <h2 className={styles.heading}>14) Documents</h2>

      <div className={styles.contentCover}>
        <div className={styles.buttonRow}>
          <Button icon={<DocumentHeaderFooter24Regular />} onClick={insertHeader}>
            Insert header
          </Button>
          <Button icon={<DocumentPageNumber24Regular />} onClick={insertPageNumber}>
            Insert page number (footer)
          </Button>
          <Button icon={<DocumentSplitHint24Regular />} onClick={insertSectionBreak}>
            Insert section break
          </Button>
          <Button icon={<TableSimple24Regular />} onClick={insertTable}>
            Insert 3x3 table
          </Button>
        </div>
      </div>
    </div>
  );
}
