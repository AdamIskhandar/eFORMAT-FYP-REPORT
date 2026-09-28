import React, { useState, useEffect } from "react";
import {
  makeStyles,
  useId,
  Dropdown,
  Option,
  Button,
  Input,
  Text,
} from "@fluentui/react-components";
import { isMarginFlagSet, applyUniformMargin } from "./Section/utils/margin";

/* global Word */

const useStyles = makeStyles({
  InsideCoverWrapper: {
    paddingLeft: "10px",
  },
  heading: {
    color: "white",
    fontSize: "20px",
  },
  label: {
    marginBottom: "10px",
  },
  contentInsideCover: {
    paddingLeft: "25px",
  },
  button: {
    marginTop: "10px",
  },
  wrapButton: {
    display: "flex",
    justifyContent: "flex-start",
  },
  wrapButtonFormat: {
    display: "flex",
    justifyContent: "flex-end",
    marginRight: "40px",
    maxWidth: "250px",
    minWidth: "250px",
  },
  status: {
    display: "block",
    marginTop: "6px",
    color: "#9fd89f",
  },
  inputRow: {
    marginBottom: "10px",
  },
});

const FORMAT_FIELDS = [
  {
    id: "titleOfReport",
    label: "Title of Report",
    bookmark: "bmTitleOfReport",
    order: 1,
    fontSize: 12,
    spacingBefore: 85,
    spacingAfter: 0,
    color: "#000000",
  },
  {
    id: "fullName",
    label: "Full Name of Student",
    bookmark: "bmFullName",
    order: 2,
    fontSize: 12,
    spacingBefore: 113.4,
    spacingAfter: 0,
    color: "#000000",
  },
  {
    id: "purpose",
    label: "Purpose of submission of Report",
    bookmark: "bmPurpose",
    order: 3,
    fontSize: 12,
    spacingBefore: 113.4,
    spacingAfter: 0,
    color: "#000000",
  },
  {
    id: "faculty",
    label: "Faculty",
    bookmark: "bmFaculty",
    order: 4,
    fontSize: 12,
    spacingBefore: 85,
    spacingAfter: 0,
    color: "#000000",
  },
  {
    id: "university",
    label: "Name of University",
    bookmark: "bmUniversity",
    order: 5,
    fontSize: 12,
    spacingBefore: 0,
    spacingAfter: 0,
    color: "#000000",
  },
  {
    id: "campus",
    label: "Name of campus where candidate is registered",
    bookmark: "bmCampus",
    order: 6,
    fontSize: 12,
    spacingBefore: 0,
    spacingAfter: 0,
    color: "#000000",
  },
  {
    id: "year",
    label: "Year in which thesis is submitted to the university for examination",
    bookmark: "bmYear",
    order: 7,
    fontSize: 12,
    spacingBefore: 56.7,
    spacingAfter: 42.5,
    color: "#000000",
  },
];

export default function TitlePage() {
  const styles = useStyles();
  const dropdownId = useId("dropdown");

  const [selectedFieldId, setSelectedFieldId] = useState(null);
  const [fieldText, setFieldText] = useState("");
  const [marginStatus, setMarginStatus] = useState("");
  const [formatStatus, setFormatStatus] = useState("");
  const [marginIsSet, setMarginIsSet] = useState(false);

  // Reflects the persisted flag on mount, so reloading the add-in doesn't
  // force a redundant "Set Margin" click if it's already been done.
  useEffect(() => {
    setMarginIsSet(isMarginFlagSet());
  }, []);

  function getFormatData(_, data) {
    const field = FORMAT_FIELDS.find((f) => f.label === data.optionValue);
    setSelectedFieldId(field ? field.id : null);
    setFormatStatus("");
  }

  // ---- Margin: single source of truth for the whole document, set once ----

  async function applyMostFrontCoverMargin() {
    if (isMarginFlagSet()) {
      setMarginStatus("The margin was set.");
      setMarginIsSet(true);
      return;
    }

    await Word.run(async (context) => {
      await applyUniformMargin(context);
    });

    setMarginStatus("Margin has been set for the whole document.");
    setMarginIsSet(true);
  }

  // ---- Format / auto-placement ----

  async function ensureSkeleton(context) {
    const marker = context.document.getBookmarkRangeOrNullObject(FORMAT_FIELDS[0].bookmark);
    await context.sync();
    if (!marker.isNullObject) return;

    const ordered = [...FORMAT_FIELDS].sort((a, b) => a.order - b.order);
    ordered.forEach((field) => {
      const paragraph = context.document.body.insertParagraph("", Word.InsertLocation.end);
      paragraph.font.name = "Arial";
      paragraph.font.bold = true;
      paragraph.font.size = field.fontSize;
      paragraph.font.allCaps = true;
      paragraph.font.color = field.color;
      paragraph.alignment = "Centered";
      paragraph.spaceBefore = field.spacingBefore;
      paragraph.spaceAfter = field.spacingAfter;
      paragraph.getRange().insertBookmark(field.bookmark);
    });

    await context.sync();
  }

  async function applyMostFrontCoverFormat() {
    if (!marginIsSet) {
      setFormatStatus("Set the margin first.");
      return;
    }
    if (!selectedFieldId || !fieldText.trim()) {
      setFormatStatus("Type the text and choose a format first.");
      return;
    }
    const field = FORMAT_FIELDS.find((f) => f.id === selectedFieldId);

    await Word.run(async (context) => {
      await ensureSkeleton(context);

      const bookmarkRange = context.document.getBookmarkRangeOrNullObject(field.bookmark);
      await context.sync();
      if (bookmarkRange.isNullObject) {
        setFormatStatus("Could not find that field's placeholder. Try again.");
        return;
      }

      bookmarkRange.insertText(fieldText.trim(), Word.InsertLocation.replace);
      await context.sync();

      const updatedRange = context.document.getBookmarkRangeOrNullObject(field.bookmark);
      updatedRange.font.name = "Arial";
      updatedRange.font.bold = true;
      updatedRange.font.size = field.fontSize;
      updatedRange.font.allCaps = true;
      updatedRange.font.color = field.color;
      updatedRange.paragraphs.load("items");
      await context.sync();

      updatedRange.paragraphs.items.forEach((p) => {
        p.alignment = "Centered";
        p.spaceBefore = field.spacingBefore;
        p.spaceAfter = field.spacingAfter;
      });
      await context.sync();

      setFormatStatus(`Inserted into: ${field.label}`);
    });

    setFieldText("");
  }

  return (
    <div className={styles.InsideCoverWrapper}>
      <h2 className={styles.heading}>1) Title Page</h2>

      <div className={styles.contentInsideCover}>
        <p className={styles.label}>Set Margin (applies once, to the whole document)</p>
        <div className={styles.wrapButton}>
          <Button shape="circular" onClick={applyMostFrontCoverMargin}>
            Set Margin
          </Button>
        </div>
        {marginStatus && <Text className={styles.status}>{marginStatus}</Text>}

        <p className={styles.label} style={{ marginTop: 16 }}>
          Type the text, then choose what it is:
        </p>
        {!marginIsSet && (
          <Text className={styles.status} style={{ color: "#e0a458" }}>
            Set the margin above first to unlock formatting.
          </Text>
        )}
        <div className={styles.inputRow}>
          <Input
            placeholder="e.g. MPP INFORMATION APP"
            value={fieldText}
            onChange={(_, data) => setFieldText(data.value)}
            style={{ width: "100%" }}
            disabled={!marginIsSet}
          />
        </div>

        <Dropdown
          id={`${dropdownId}-format`}
          placeholder="Select the format"
          value={selectedFieldId ? FORMAT_FIELDS.find((f) => f.id === selectedFieldId).label : ""}
          onOptionSelect={getFormatData}
          disabled={!marginIsSet}
        >
          {FORMAT_FIELDS.map((field) => (
            <Option key={field.id} value={field.label}>
              {field.label}
            </Option>
          ))}
        </Dropdown>
        <div className={styles.wrapButtonFormat}>
          <Button
            className={styles.button}
            shape="circular"
            onClick={applyMostFrontCoverFormat}
            disabled={!marginIsSet}
          >
            Set Format
          </Button>
        </div>
        {formatStatus && <Text className={styles.status}>{formatStatus}</Text>}
      </div>
    </div>
  );
}
