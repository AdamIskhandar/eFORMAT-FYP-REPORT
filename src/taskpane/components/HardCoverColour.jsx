import React from "react";
import { makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
  HardCoverWrapper: {
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
});

export default function HardCoverColour() {
  const styles = useStyles();
  return (
    <>
      <div className={styles.HardCoverWrapper}>
        <h2 className={styles.heading}>Hard Cover Colour</h2>
        <div>
          <p className={styles.paragprahp}>Dark Blue</p>
        </div>
      </div>
    </>
  );
}
