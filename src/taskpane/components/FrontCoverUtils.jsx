import React, { useState, useEffect } from "react";
import { Button, Dropdown, Option, Input, Field } from "@fluentui/react-components";
import { CheckmarkCircle20Filled, DocumentAdd20Regular } from "@fluentui/react-icons";
import { generateFrontCover, frontCoverExists, fitCoverOnOnePage } from "./Section/utils/front-cover-utils";
import { setFillInText } from "./Section/utils/fill-in-fields";
import { StepBody, SubHeading, Group, Actions, Status, GuideTable, GuideSteps } from "./ui/StepUI";

/* global Word */

// Guide table only - the real values used when generating live are in
// front-cover-utils.js. Kept here so students can see the spec at a glance.
const itemsMostFrontCover = [
  { element: "Top / Bottom Margin", format: "4 cm" },
  { element: "Left / Right Margin", format: "3.2 cm" },
  { element: "UNIVERSITI POLY-TECH MALAYSIA", format: "Arial Bold 22 - first line (fixed)" },
  { element: "Project Title", format: "Arial Bold 20 - centre of page" },
  { element: "Student Full Name", format: "Arial Bold 20 - last line" },
];

// Lines the student can fill from the panel (tags match front-cover-utils.js).
const COVER_FIELDS = [
  { tag: "eformat-cover-title", bookmark: "bmFcProjectTitle", label: "Project Title", size: 20, example: "e.g. MPP INFORMATION APP" },
  { tag: "eformat-cover-name", bookmark: "bmFcFullName", label: "Student Full Name", size: 20, example: "e.g. ADAM ISKHANDAR BIN MUHASLEE" },
];

export default function FrontCoverUtils() {
  const [status, setStatus] = useState("");
  const [statusTone, setStatusTone] = useState("success");
  const [alreadyDone, setAlreadyDone] = useState(false);
  const [fieldTag, setFieldTag] = useState("");
  const [fieldText, setFieldText] = useState("");
  const [fillStatus, setFillStatus] = useState("");

  useEffect(() => {
    // Check the document itself (not a saved flag), so the button only says
    // "Generated" when the cover is really there.
    Word.run(async (context) => {
      setAlreadyDone(await frontCoverExists(context));
    }).catch(() => setAlreadyDone(false));
  }, []);

  async function handleGenerate() {
    try {
      await Word.run(async (context) => {
        const { created } = await generateFrontCover(context);
        setStatusTone("success");
        setStatus(
          created
            ? "Front Cover generated. Now fill in the red lines using Fill in the cover below: choose a line, type your text, click Set Text."
            : "Front Cover was already generated. Use Fill in the cover below to change its text."
        );
        setAlreadyDone(true);
      });
    } catch (err) {
      console.error(err);
      setStatusTone("error");
      setStatus(err.message || "Could not generate the Front Cover.");
    }
  }

  async function handleFill() {
    const field = COVER_FIELDS.find((f) => f.tag === fieldTag);
    if (!field || !fieldText.trim()) {
      setFillStatus("Choose the line and type the text first.");
      return;
    }
    try {
      await Word.run(async (context) => {
        const ok = await setFillInText(context, field, fieldText.trim(), field.size);
        if (!ok) {
          setFillStatus("Line not found. Generate the Front Cover in a new document first.");
          return;
        }
        // Keep the cover on one page: long text -> smaller gaps.
        const spacing = await fitCoverOnOnePage(context);
        let note = "";
        if (spacing && !spacing.fits) {
          note = " This text is too long to fit on one page even with smaller gaps - please shorten it.";
        } else if (spacing && spacing.shrunk) {
          note = " The spacing was made smaller so the cover stays on one page.";
        }
        setFillStatus(`Updated: ${field.label}.${note}`);
        setFieldText("");
      });
    } catch (err) {
      console.error(err);
      setFillStatus(err.message || "Could not update the cover.");
    }
  }
  const fillTone = /shorten it\.$/.test(fillStatus)
    ? "warning"
    : fillStatus.startsWith("Updated")
    ? "success"
    : /first\.$/.test(fillStatus)
    ? "warning"
    : "error";
  const chosen = COVER_FIELDS.find((f) => f.tag === fieldTag);

  const tone = statusTone;

  return (
    <StepBody>
      <GuideSteps
        steps={[
          <>Click <b>Generate Front Cover</b> (only once).</>,
          <>In <b>Fill in the cover</b> below, choose a line: Project Title or Student Full Name.</>,
          <>Type your text in the box.</>,
          <>Click <b>Set Text</b>. The red words on the cover change to your text in black.</>,
          <>Do steps 2 - 4 again for the other line.</>,
        ]}
        note="Fill in the cover using the boxes here - you don't need to type on the cover in Word."
      />
      <Actions stretch>
        <Button
          appearance={alreadyDone ? "secondary" : "primary"}
          icon={alreadyDone ? <CheckmarkCircle20Filled /> : <DocumentAdd20Regular />}
          onClick={handleGenerate}
        >
          {alreadyDone ? "Front Cover Generated" : "Generate Front Cover"}
        </Button>
      </Actions>
      <Status text={status} tone={tone} />

      <Group>
        <SubHeading>Fill in the cover</SubHeading>
        <Field label="1. Choose the line">
          <Dropdown
            placeholder="Select the line to fill"
            value={chosen ? chosen.label : ""}
            selectedOptions={fieldTag ? [fieldTag] : []}
            onOptionSelect={(_, data) => {
              setFieldTag(data.optionValue);
              setFillStatus("");
            }}
          >
            {COVER_FIELDS.map((f) => (
              <Option key={f.tag} value={f.tag}>
                {f.label}
              </Option>
            ))}
          </Dropdown>
        </Field>
        <Field label="2. Type the text">
          <Input
            placeholder={chosen ? chosen.example : "Type here"}
            value={fieldText}
            onChange={(_, data) => setFieldText(data.value)}
            onKeyDown={(e) => e.key === "Enter" && handleFill()}
          />
        </Field>
        <Actions>
          <Button appearance="primary" onClick={handleFill}>
            Set Text
          </Button>
        </Actions>
        <Status text={fillStatus} tone={fillTone} />
      </Group>

      <GuideTable title="Most-front Cover format" rows={itemsMostFrontCover} />
    </StepBody>
  );
}
