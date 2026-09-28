import React, { useId, useState } from "react";
import { Button, Dropdown, makeStyles, Option } from "@fluentui/react-components";

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

function ListOfTables() {
  const dropdownId = useId("dropdown");

  // for text
  const textOptions = ["Title of Table"];
  const [text, setText] = useState("null");

  // function to get and set dropdown data for margin
  async function getTextData(e, data) {
    setText(data.optionValue);
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

  async function applyTextFormat() {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const paragraphs = selection.paragraphs;
      const sections = selection.sections;
      // const body = context.document.body;
      selection.load("text");
      paragraphs.load("items");
      sections.load("items");
      await context.sync();

      paragraphs.items.forEach((p) => {
        if (text === "Title of Table") {
          if (sections.items.length === 0) {
            console.log("No section found for this selection.");
            return;
          }

          const currentSection = sections.items[0];

          // // Step 4: Conversion helper (cm → points)
          const inchToPt = (inch) => inch * 72;

          // // Step 5: Apply custom margin to this section only
          currentSection.pageSetup.topMargin = inchToPt(1.18);
          p.font.bold = true;
          p.alignment = "Centered";
          p.font.size = 12;
          p.font.name = "Arial";
        }
      });

      setText("null");

      await context.sync();
    });
  }

  const styles = useStyles();
  return (
    <>
      <div className={styles.ContentWrapper}>
        <h2 className={styles.heading}>6) List Of Tables</h2>
        <div className={styles.contentCover}>
          {/* <p className={styles.label}>Select The Margin</p>
          <div className={styles.wrapButton}>
            <Button shape="circular" onClick={applyContentMargin}>
              Set Margin
            </Button>
          </div> */}

          <p className={styles.label}>Select format for title :</p>
          <Dropdown
            id={`${dropdownId}-table`}
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
        </div>
      </div>
    </>
  );
}

export default ListOfTables;
