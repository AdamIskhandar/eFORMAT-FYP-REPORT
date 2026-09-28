import React, { useState } from "react";
import {
  Button,
  Dropdown,
  Option,
  Field,
  Toolbar,
  ToolbarButton,
  ToolbarDivider,
  makeStyles,
} from "@fluentui/react-components";
import {
  TextBold24Regular,
  TextItalic24Regular,
  TextUnderline24Regular,
  ColorLine24Regular,
} from "@fluentui/react-icons";

/* global Word */

const FONTS = ["Calibri", "Arial", "Times New Roman", "Georgia", "Verdana"];
const SIZES = [9, 10, 11, 12, 14, 16, 18, 24, 32];

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
});

export default function FormattingEngine() {
  const styles = useStyles();
  const [font, setFont] = useState("Calibri");
  const [size, setSize] = useState(11);

  async function applyToSelection(mutator) {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      mutator(range.font);
      await context.sync();
    });
  }

  async function toggleBoldReal() {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.font.load("bold");
      await context.sync();
      range.font.bold = !range.font.bold;
      await context.sync();
    });
  }

  async function toggleItalicReal() {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.font.load("italic");
      await context.sync();
      range.font.italic = !range.font.italic;
      await context.sync();
    });
  }

  async function toggleUnderlineReal() {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.font.load("underline");
      await context.sync();
      range.font.underline =
        range.font.underline === Word.UnderlineType.none
          ? Word.UnderlineType.single
          : Word.UnderlineType.none;
      await context.sync();
    });
  }

  async function applyFont(nextFont) {
    setFont(nextFont);
    await applyToSelection((f) => (f.name = nextFont));
  }

  async function applySize(nextSize) {
    setSize(nextSize);
    await applyToSelection((f) => (f.size = nextSize));
  }

  async function applyColor(hex) {
    await applyToSelection((f) => (f.color = hex));
  }

  return (
    <div className={styles.ContentWrapper}>
      <h2 className={styles.heading}>13) Formatting</h2>

      <div className={styles.contentCover}>
        <Toolbar aria-label="Text formatting">
          <ToolbarButton icon={<TextBold24Regular />} onClick={toggleBoldReal}>
            Bold
          </ToolbarButton>
          <ToolbarButton icon={<TextItalic24Regular />} onClick={toggleItalicReal}>
            Italic
          </ToolbarButton>
          <ToolbarButton icon={<TextUnderline24Regular />} onClick={toggleUnderlineReal}>
            Underline
          </ToolbarButton>
          <ToolbarDivider />
          <ToolbarButton icon={<ColorLine24Regular />} onClick={() => applyColor("#D13438")}>
            Red
          </ToolbarButton>
          <ToolbarButton icon={<ColorLine24Regular />} onClick={() => applyColor("#000000")}>
            Black
          </ToolbarButton>
        </Toolbar>

        <div className={styles.label}>
          <Field label="Font family">
            <Dropdown value={font} onOptionSelect={(_, data) => applyFont(data.optionValue)}>
              {FONTS.map((f) => (
                <Option key={f} value={f}>
                  {f}
                </Option>
              ))}
            </Dropdown>
          </Field>
        </div>

        <div className={styles.label}>
          <Field label="Font size">
            <Dropdown
              value={String(size)}
              onOptionSelect={(_, data) => applySize(Number(data.optionValue))}
            >
              {SIZES.map((s) => (
                <Option key={s} value={String(s)}>
                  {s}
                </Option>
              ))}
            </Dropdown>
          </Field>
        </div>
      </div>
    </div>
  );
}
