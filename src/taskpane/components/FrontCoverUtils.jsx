import React, { useState, useEffect } from "react";
import { Button } from "@fluentui/react-components";
import { CheckmarkCircle20Filled, DocumentAdd20Regular } from "@fluentui/react-icons";
import { generateFrontCover, isFrontCoverGenerated } from "./Section/utils/front-cover-utils";
import { StepBody, Hint, Actions, Status, GuideTable } from "./ui/StepUI";

/* global Word */

// Guide table only - the real values used when generating live are in
// front-cover-utils.js. Kept here so students can see the spec at a glance.
const itemsMostFrontCover = [
  { element: "Top Margin", format: "4 cm" },
  { element: "Bottom Margin", format: "4 cm" },
  { element: "Left Margin", format: "3.2 cm" },
  { element: "Right Margin", format: "3.2 cm" },
  { element: "UNIVERSITI POLY-TECH MALAYSIA", format: "Arial Bold 20 - First Line on page" },
  { element: "Project Title", format: "Arial Bold 20 - Centre of Page" },
  { element: "Student's Full Name", format: "Arial Bold 20 - Last line on page (red)" },
  {
    element: "Programme (e.g. Computer Application Development)",
    format: "Arial Bold 20 - last line, specialisation in red",
  },
];

export default function FrontCoverUtils() {
  const [status, setStatus] = useState("");
  const [alreadyDone, setAlreadyDone] = useState(false);

  useEffect(() => {
    setAlreadyDone(isFrontCoverGenerated());
  }, []);

  async function handleGenerate() {
    try {
      await Word.run(async (context) => {
        const { created } = await generateFrontCover(context);
        setStatus(
          created
            ? "Most-front Cover generated above the Title Page. Replace the [bracketed] placeholders with your own text."
            : "Most-front Cover was already generated - edit the bracketed placeholders directly in the document."
        );
        setAlreadyDone(true);
      });
    } catch (err) {
      console.error(err);
      setStatus(err.message || "Could not generate the Front Cover.");
    }
  }

  // Success messages start with "Most-front Cover"; anything else is an error.
  const tone = status.startsWith("Most-front Cover") ? "success" : "error";

  return (
    <StepBody>
      <Hint>
        Start here. Creates the most-front cover on A4 paper with its own 4 cm / 3.2 cm margin,
        placed above the Title Page, and sets up the sections the next steps use. Runs once.
      </Hint>
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
      <GuideTable title="Most-front Cover format" rows={itemsMostFrontCover} />
    </StepBody>
  );
}
