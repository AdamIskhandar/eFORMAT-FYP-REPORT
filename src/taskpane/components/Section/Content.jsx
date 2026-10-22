import React, { useState } from "react";
import { useId, Dropdown, Option, Button, Field } from "@fluentui/react-components";
import { TextGrammarWand20Regular, DocumentText20Regular } from "@fluentui/react-icons";
import { StepBody, SubHeading, Group, Actions, Status, GuideTable, GuideSteps } from "../ui/StepUI";
import { CONTENT_FORMAT, formatSelection, formatWholeReport } from "./utils/content-format-utils";

/* global Word */

// Guide only - the real values are CONTENT_FORMAT in content-format-utils.js.
const GUIDE = [
  { element: "Font/Size (Heading)", format: "Arial / 16-point / Bold" },
  { element: "Font/Size (Sub heading)", format: "Arial / 14-point / Bold" },
  { element: "Font/Size (Content)", format: "Arial / 11-point" },
  { element: "Font/Size (in table) / Label Figure or Table", format: "Arial / 10-point" },
  { element: "Spacing", format: "1.5" },
  { element: "Alignment", format: "Justify" },
];

const KINDS = Object.entries(CONTENT_FORMAT).map(([key, f]) => ({ key, label: f.label }));

export default function Content() {
  const dropdownId = useId("content-kind");

  const [kind, setKind] = useState("");
  const [selMsg, setSelMsg] = useState({ text: "", isError: false });
  const [allMsg, setAllMsg] = useState({ text: "", isError: false });

  async function handleFormatSelection() {
    if (!kind) {
      setSelMsg({ text: "Choose what the selected text is first.", isError: true });
      return;
    }
    try {
      await Word.run(async (context) => {
        const n = await formatSelection(context, kind);
        setSelMsg({
          text: `Formatted ${n} paragraph${n === 1 ? "" : "s"} as ${CONTENT_FORMAT[kind].label}.`,
          isError: false,
        });
      });
    } catch (err) {
      console.error(err);
      setSelMsg({ text: err.message || "Could not format the selection.", isError: true });
    }
  }

  async function handleFormatAll() {
    try {
      await Word.run(async (context) => {
        const { content, small } = await formatWholeReport(context);
        setAllMsg({
          text: `Formatted ${content} content paragraph${content === 1 ? "" : "s"} and ${small} table/caption line${small === 1 ? "" : "s"}.`,
          isError: false,
        });
      });
    } catch (err) {
      console.error(err);
      setAllMsg({ text: err.message || "Could not format the report.", isError: true });
    }
  }

  const msg = (m) => <Status text={m.text} tone={m.isError ? "error" : "success"} />;

  return (
    <StepBody>
      <GuideSteps
        steps={[
          <>Write your report first (from Chapter 1 to the end).</>,
          <>Click <b>Format Whole Report</b>. Text becomes Arial 11, 1.5 spacing, justify; text in tables and Figure/Table captions becomes Arial 10.</>,
          <>To fix only one part: in Word, click in a paragraph (or select several), choose what it is in <b>The selected text is</b>, then click <b>Set Format</b>.</>,
        ]}
        note="Headings and the pages before Chapter 1 are not changed. Bold or italic words you set yourself are kept."
      />
      <GuideTable title="UPTM content format" rows={GUIDE} />

      <Group>
        <SubHeading>Format the whole report</SubHeading>
        <Actions stretch>
          <Button appearance="primary" icon={<DocumentText20Regular />} onClick={handleFormatAll}>
            Format Whole Report
          </Button>
        </Actions>
        {msg(allMsg)}
      </Group>

      <Group>
        <SubHeading>Format selected text</SubHeading>
        <Field label="The selected text is">
          <Dropdown
            id={dropdownId}
            placeholder="Select the element"
            value={kind ? CONTENT_FORMAT[kind].label : ""}
            selectedOptions={kind ? [kind] : []}
            onOptionSelect={(_, data) => setKind(data.optionValue)}
          >
            {KINDS.map((k) => (
              <Option key={k.key} value={k.key}>
                {k.label}
              </Option>
            ))}
          </Dropdown>
        </Field>
        <Actions>
          <Button appearance="primary" icon={<TextGrammarWand20Regular />} onClick={handleFormatSelection}>
            Set Format
          </Button>
        </Actions>
        {msg(selMsg)}
      </Group>
    </StepBody>
  );
}
