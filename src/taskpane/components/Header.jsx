import * as React from "react";
import { tokens, makeStyles, Text } from "@fluentui/react-components";
import backgroundImage from "../../../assets/banner.png";

const useStyles = makeStyles({
  header: {
    display: "flex",
    flexDirection: "column",
    gap: "10px",
    padding: "12px 12px 6px 12px",
  },
  banner: {
    width: "100%",
    height: "130px",
    borderRadius: tokens.borderRadiusXLarge,
    backgroundImage: `url(${backgroundImage})`,
    backgroundRepeat: "no-repeat",
    backgroundPosition: "center",
    backgroundSize: "cover",
    boxShadow: tokens.shadow8,
  },
  titles: {
    display: "flex",
    flexDirection: "column",
    padding: "0 4px",
  },
  title: {
    fontSize: tokens.fontSizeBase500,
    fontWeight: tokens.fontWeightBold,
    color: tokens.colorNeutralForeground1,
  },
  subtitle: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
});

const Header = () => {
  const styles = useStyles();
  return (
    <header className={styles.header}>
      <div className={styles.banner} role="img" aria-label="UPTM eFormat FYP Report" />
      <div className={styles.titles}>
        <Text className={styles.title}>Report FYP Template TESTTT</Text>
        <Text className={styles.subtitle}>Formats your FYP report to the UPTM guidelines</Text>
      </div>
    </header>
  );
};

export default Header;
