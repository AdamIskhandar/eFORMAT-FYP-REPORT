import React from "react";
import { makeStyles, tokens, Text } from "@fluentui/react-components";
import { StepBody, SubHeading } from "./ui/StepUI";

const useStyles = makeStyles({
  row: {
    display: "flex",
    alignItems: "center",
    gap: "12px",
  },
  swatch: {
    width: "36px",
    height: "36px",
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: "#0B1F4B",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    flexShrink: 0,
  },
  name: {
    fontWeight: tokens.fontWeightSemibold,
  },
});

export default function HardCoverColour() {
  const styles = useStyles();
  return (
    <StepBody>
      <SubHeading>Hard Cover Colour</SubHeading>
      <div className={styles.row}>
        <span className={styles.swatch} aria-hidden="true" />
        <Text className={styles.name}>Dark Blue</Text>
      </div>
    </StepBody>
  );
}
