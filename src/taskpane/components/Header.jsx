import * as React from "react";
import PropTypes from "prop-types";
import { Image, tokens, makeStyles } from "@fluentui/react-components";
import backgroundImage from "../../../assets/uptm3.jpg";

const useStyles = makeStyles({
  welcome__header: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    paddingTop: "20px",
    paddingLeft: "10px",
    paddingRight: "10px",
    position: "relative",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
  },
  background: {
    position: "absolute",
    top: 0,
    left: 0,
    width: "100%",
    height: "100%",
    backgroundImage: `url(${backgroundImage})`,
    backgroundRepeat: "no-repeat",
    backgroundPosition: "center",
    backgroundSize: "cover",
    filter: "blur(4px)",
    backgroundColor: "rgba(0,0,0,0.1)",
    backgroundBlendMode: "darken",
    zIndex: 0,
  },
  message: {
    fontSize: "clamp(1.2rem, 2vw + 0.5rem, 1.5rem)",
    fontWeight: tokens.fontWeightRegular,
    color: "white",
    textAlign: "center",
    fontWeight: tokens.fontWeightBold,
    paddingTop: "15px",
    // backgroundColor: "rgba(0, 0, 0, 0.5)",
    // borderRadius: "8px",
    // backdropFilter: "blur(2px)",
    zIndex: 3,
  },

  image: {
    zIndex: 2,
  },
});

// improvementtss

// 1) submit report via add ins - RM 3k
// 2) submit progress report via add ins - RM 4k
// 3) sign page for security - RM 5k

const Header = (props) => {
  const { title, logo, message } = props;
  const styles = useStyles();

  return (
    <section className={styles.welcome__header}>
      <div className={styles.background}></div>
      <Image width="300" src={logo} alt={title} className={styles.image} />
      <h1 className={styles.message}>{message}</h1>
    </section>
  );
};

export default Header;
