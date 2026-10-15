import React from "react";
import { StepBody, SubHeading, GuideTable } from "../ui/StepUI";

// Guide only - the side cover (spine) is printed by the binder.
const ROWS = [
  { element: "Student ID", format: "Arial 11" },
  { element: "Category", format: "Arial 11" },
  { element: "Year", format: "Arial 11" },
];

export default function SideCover() {
  return (
    <StepBody>
      <SubHeading>Side Cover</SubHeading>
      <GuideTable rows={ROWS} />
    </StepBody>
  );
}
