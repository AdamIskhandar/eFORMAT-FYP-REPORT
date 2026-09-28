import React from "react";
import { makeStyles, useId, Dropdown, Option, Button } from "@fluentui/react-components";
import { useState } from "react";

const useStyles = makeStyles({
  ContentWrapper: {
    paddingLeft: "10px",
  },

  heading: {
    color: "white",
    fontSize: "20px",
  },

  paragprahp: {
    paddingLeft: "25px",
    fontSize: "15px",
  },

  label: {
    marginBottom: "10px",
  },

  contentCover: {
    paddingLeft: "25px",
  },

  button: {
    marginTop: "10px",
  },

  wrapButton: {
    display: "flex",
    justifyContent: "flex-start",
  },

  wrapButtonSize: {
    display: "flex",
    justifyContent: "flex-end",
    marginRight: "40px",
    maxWidth: "250px",
    minWidth: "250px",
  },
});

export default function Content() {
  const styles = useStyles();
  const dropdownId = useId("dropdown");

  // for text
  const textOptions = ["Text In Table/Label Table", "Figure Text"];
  const [text, setText] = useState("null");

  // for fontSize
  const sizeOptions = ["Heading", "SubHeading", "Content", "Spacing", "Alignment (Justify)"];
  const [size, setSize] = useState("null");

  // function to get and set dropdown data for margin
  async function getTextData(e, data) {
    setText(data.optionValue);
  }

  // function to get and aet dropdown data for size
  async function getSizeData(e, data) {
    setSize(data.optionValue);
  }

  // function ni RM10k
  async function divideSection() {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertBreak(Word.BreakType.sectionNext, Word.InsertLocation.end);

      await context.sync();

      console.log("Inserted section break on next page.");
    });
  }

  // function to apply the margin into the WORD
  async function applyContentMargin() {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();

      const sections = selection.sections;
      sections.load("items");
      await context.sync();

      // buat function untuk bahagi section

      // function tu kena auto bahagikan section untuk setiap page
      // function call dari button

      if (sections.items.length === 0) {
        console.log("No section found for this selection.");
        return;
      }

      const currentSection = sections.items[0];

      // // Step 4: Conversion helper (cm → points)
      const inchToPt = (inch) => inch * 72;

      // // Step 5: Apply custom margin to this section only
      currentSection.pageSetup.topMargin = inchToPt(1);
      currentSection.pageSetup.bottomMargin = inchToPt(1.2);
      currentSection.pageSetup.leftMargin = inchToPt(1.52);
      currentSection.pageSetup.rightMargin = inchToPt(1.2);

      await context.sync();
    });
  }

  // function to apply the size into the WORD
  async function applySizeFormat() {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const paragraphs = selection.paragraphs;
      // const body = context.document.body;
      selection.load("text");
      paragraphs.load("items");
      await context.sync();

      paragraphs.items.forEach((p) => {
        if (size === "Heading") {
          p.font.name = "Arial";
          p.font.bold = true;
          p.font.size = 16;
        } else if (size === "SubHeading") {
          p.font.name = "Arial";
          p.font.bold = true;
          p.font.size = 14;
        } else if (size === "Content") {
          p.font.name = "Arial";
          p.font.bold = false;
          p.font.size = 11;
        } else if (size === "Spacing") {
          p.lineSpacing = 18;
        } else if (size === "Alignment (Justify)") {
          p.alignment = "Justified";
        }
      });

      setSize("null");

      await context.sync();
    });
  }

  // function to apply the text format into the WORD
  async function applyTextFormat() {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const paragraphs = selection.paragraphs;
      // const body = context.document.body;
      selection.load("text");
      paragraphs.load("items");
      await context.sync();

      paragraphs.items.forEach((p) => {
        if (text === "Text In Table/Label Table") {
          p.font.name = "Arial";
          p.font.bold = true;
          p.font.size = 10;
          p.lineSpacing = 12;
        } else if (text === "Figure Text") {
          p.font.name = "Arial";
          p.font.bold = true;
          p.font.size = 10;
          p.lineSpacing = 12;
        }
      });

      setText("null");

      await context.sync();
    });
  }

  return (
    <>
      <div className={styles.ContentWrapper}>
        <h2 className={styles.heading}>5) Content</h2>

        {/* for margin */}
        <div className={styles.contentCover}>
          <p className={styles.label}>Select The Margin</p>
          <div className={styles.wrapButton}>
            <Button shape="circular" onClick={applyContentMargin}>
              Set Margin
            </Button>
          </div>

          {/* for format  */}
          <p className={styles.label}>Select size for :</p>
          <Dropdown
            id={`${dropdownId}-size`}
            placeholder="Select the element size"
            value={size === "null" ? "" : size}
            onOptionSelect={getSizeData}
          >
            {sizeOptions.map((option) => (
              <Option key={option}>{option}</Option>
            ))}
          </Dropdown>
          <div className={styles.wrapButtonSize}>
            <Button className={styles.button} shape="circular" onClick={applySizeFormat}>
              Set Size
            </Button>
          </div>

          {/* for text in table/label table or figure text in table */}
          <p className={styles.label}>Select format for text (table/figure) :</p>
          <Dropdown
            id={`${dropdownId}-text`}
            placeholder="Select the text"
            value={text === "null" ? "" : text}
            onOptionSelect={getTextData}
          >
            {textOptions.map((option) => (
              <Option key={option}>{option}</Option>
            ))}
          </Dropdown>
          <div className={styles.wrapButtonSize}>
            <Button className={styles.button} shape="circular" onClick={applyTextFormat}>
              Set Text Format
            </Button>
          </div>
          {/* <div>
            <Button className={styles.button} shape="circular" onClick={divideSection}>
              divide section
            </Button>
          </div> */}
        </div>
      </div>
    </>
  );
}
