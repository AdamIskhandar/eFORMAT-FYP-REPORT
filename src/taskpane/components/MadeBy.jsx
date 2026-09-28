import React from "react";
import { makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
  madeByWrapper: {
    display: "flex",
    justifyContent: "center",
    fontWeight: "bold",
  },
});

export default function MadeBy() {
  const styles = useStyles();

  return (
    <>
      <div className={styles.madeByWrapper}>
        <p>Made By Adam Iskhandar 👋✨</p>
      </div>
    </>
  );
}
