import * as React from "react";
import Header from "./Header";

import { makeStyles, tokens } from "@fluentui/react-components";
import HardCoverColour from "./HardCoverColour";
import Content from "./Section/Content";
import SideCover from "./Section/SideCover";
import MadeBy from "./MadeBy";
import Declaration from "./Section/Declaration";
import Acknowledgements from "./Section/Acknowledgements";
import Abstract from "./Section/Abstract";
import ListOfTables from "./Section/ListOfTables";
import References from "./Section/References";
import Appendices from "./Section/Appendices";
import CaptionCrossRefPanel from "./Section/CaptionCrossRefPanel";
import CitationReferencePanel from "./Section/CitationPanel";
import FormattingEngine from "./Section/FormattingEngine";
import DocumentModules from "./Section/DocumentModules";
import ReferenceModules from "./Section/ReferenceModules";
import ConfigSettings from "./Section/ConfigSettings";
import DashboardPanel from "./Section/DashboardPanel";
import FrontCover from "./FrontCover";
import TitlePage from "./TitlePage";
import CaptionLists from "./Section/CaptionLists";
import TableOfContents from "./Section/TableOfContents";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    fontFamily: tokens.fontFamilyBase,
  },
  header: {
    padding: "12px 16px 4px 16px",
  },
  panel: {
    flex: 1,
    overflowY: "auto",
    padding: "12px 16px",
  },
});

// const MENU = [
//   { value: "formatting", label: "Formatting", icon: <TextBold24Regular /> },
//   { value: "document", label: "Document", icon: <DocumentPageNumber24Regular /> },
//   { value: "reference", label: "Reference", icon: <BookOpen24Regular /> },
//   { value: "settings", label: "Settings", icon: <Settings24Regular /> },
// ];

export default function App(props) {
  const styles = useStyles();
  // const [selected, setSelected] = useState("formatting");

  return (
    <div className={styles.root}>
      <Header logo="assets/uptmlogo.png" title={props.title} message="UPTM REPORT FORMATTING" />
      {/* <TextInsertion insertText={insertText} /> */}
      {/* Hard cover colour ✅ */}
      <HardCoverColour />
      {/* Most-front Cover (hard binding) ✅*/}
      <FrontCover />
      {/* Side Cover ✅*/}
      <SideCover />
      {/* title page ✅ */}
      <TitlePage />
      {/* DECLARATION ✅*/}
      <Declaration />
      {/* ACKNOWLEDGEMENTS ✅ */}
      <Acknowledgements />
      {/* ABSTRACT ✅ */}
      <Abstract />
      {/* TABLE OF CONTENT ✅ */}
      <TableOfContents />

      {/* LIST OF TABLES ✅ */}
      {/* <ListOfTables /> */}

      {/* CITATION PANEL ✅*/}
      {/* <CitationReferencePanel /> */}
      <CaptionLists />

      {/* CAPTION PANEL ✅*/}
      <CaptionCrossRefPanel />

      {/* CONTENTS ✅ */}
      <Content />

      {/* REFERENCES ✅*/}
      <References />

      {/* APPENDICES ✅*/}
      <Appendices />

      {/* FORMATTING ✅*/}
      {/* <FormattingEngine /> */}

      {/* DOCUMENT MODULES ✅*/}
      {/* <DocumentModules /> */}

      {/* REFERENCES MODULE ✅*/}
      {/* <ReferenceModules /> */}

      {/* CONFIG SETTING ✅*/}
      {/* <ConfigSettings /> */}

      {/* DASHBOARD ✅*/}
      {/* <DashboardPanel /> */}

      {/* Made By ✅*/}
      {/* <MadeBy /> */}
    </div>
  );
}
