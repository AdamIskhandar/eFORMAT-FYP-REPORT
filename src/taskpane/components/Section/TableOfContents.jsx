import React, { useState } from "react";
import { makeStyles, Button, Text } from "@fluentui/react-components";
import { generateTableOfContents } from "../Section/utils/toc-generator-utils";
import { REPORT_STRUCTURE } from "../Section/utils/toc-template-config";

/* global Word */

const useStyles = makeStyles({
  ContentWrapper: {
    paddingLeft: "10px",
  },
  heading: {
    color: "white",
    fontSize: "20px",
  },
  Wrapper: {
    paddingLeft: "25px",
    display: "flex",
    flexDirection: "column",
    gap: "10px",
    maxWidth: "320px",
  },
  note: {
    color: "#c7c7c7",
  },
  status: {
    display: "block",
    marginTop: "6px",
    color: "#9fd89f",
  },
});

export default function TableOfContents() {
  const styles = useStyles();
  const [status, setStatus] = useState("");

  async function handleGenerate() {
    await Word.run(async (context) => {
      const { created, count } = await generateTableOfContents(context, REPORT_STRUCTURE);
      setStatus(
        created
          ? `Table of Contents generated with ${count} entries - each chapter/subsection is now a real heading further down the document. Right-click the page numbers and choose Update Field once your report has taken shape.`
          : "Table of Contents was already generated. Edit the headings directly in the document from here on - clicking Generate again won't rebuild it, so your content stays safe."
      );
    });
  }

  return (
    <div className={styles.ContentWrapper}>
      <h2 className={styles.heading}>5) Table of Contents</h2>
      <div className={styles.Wrapper}>
        <Text className={styles.note}>
          Generates the Table of Contents page and creates every chapter and subsection as a real
          heading in the document, with a placeholder telling you what to write. This only runs
          once - it won't overwrite anything you've already written.
        </Text>

        <Button appearance="primary" onClick={handleGenerate}>
          Generate Table of Contents
        </Button>

        {status && <Text className={styles.status}>{status}</Text>}
      </div>
    </div>
  );
}
