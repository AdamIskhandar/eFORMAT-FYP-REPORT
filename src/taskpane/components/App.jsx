import * as React from "react";
import {
  makeStyles,
  mergeClasses,
  tokens,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  Button,
  Text,
} from "@fluentui/react-components";
import Header from "./Header";
import WhatsNew from "./WhatsNew";
import FrontCoverUtils from "./FrontCoverUtils";
import TitlePage from "./TitlePage";
import HeaderFooter from "./Section/HeaderFooter";
import Declaration from "./Section/Declaration";
import Acknowledgements from "./Section/Acknowledgements";
import Abstract from "./Section/Abstract";
import TableOfContents from "./Section/TableOfContents";
import CaptionLists from "./Section/CaptionLists";
import CaptionCrossRefPanel from "./Section/CaptionCrossRefPanel";
import Content from "./Section/Content";

// Not shown in the panel right now (kept for later):
// HardCoverColour, SideCover (Printing & Binding guide), FrontCover, ListOfTables, CitationPanel, References,
// Appendices, FormattingEngine, DocumentModules, ReferenceModules,
// ConfigSettings, DashboardPanel, MadeBy

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    minHeight: "100vh",
    fontFamily: tokens.fontFamilyBase,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  intro: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "8px",
    padding: "4px 12px 8px 12px",
  },
  introText: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  groupLabel: {
    display: "block",
    padding: "14px 16px 6px 16px",
    fontSize: tokens.fontSizeBase100,
    fontWeight: tokens.fontWeightSemibold,
    letterSpacing: "0.08em",
    textTransform: "uppercase",
    color: tokens.colorBrandForeground1,
  },
  list: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    padding: "0 10px",
  },
  item: {
    borderRadius: tokens.borderRadiusLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    overflow: "hidden",
  },
  itemOpen: {
    border: `1px solid ${tokens.colorBrandStroke1}`,
  },
  headerRow: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    textAlign: "left",
    padding: "2px 0",
  },
  badge: {
    flexShrink: 0,
    width: "24px",
    height: "24px",
    borderRadius: "50%",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightBold,
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  badgeInfo: {
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground2,
  },
  titles: {
    display: "flex",
    flexDirection: "column",
    minWidth: 0,
  },
  title: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  subtitle: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  panel: {
    padding: "0 12px",
  },
  footer: {
    padding: "18px 16px 24px 16px",
    textAlign: "center",
    color: tokens.colorNeutralForeground4,
    fontSize: tokens.fontSizeBase100,
  },
});

// The steps, in the order a student does them. Step numbers come from this
// order, so they always stay right when steps are added or moved.
const GROUPS = [
  {
    label: "Cover",
    steps: [
      { id: "front-cover", title: "Front Cover", subtitle: "Creates the cover, A4 paper and sections", Component: FrontCoverUtils },
      { id: "title-page", title: "Title Page", subtitle: "Report margin and title page text", Component: TitlePage },
      { id: "header-footer", title: "Header & Footer", subtitle: "Name, title, subject code, page X of Y", Component: HeaderFooter },
    ],
  },
  {
    label: "Front pages",
    steps: [
      { id: "declaration", title: "Declaration", subtitle: "Date, name and student ID", Component: Declaration },
      { id: "acknowledgements", title: "Acknowledgements", subtitle: "Formatted page with sample text", Component: Acknowledgements },
      { id: "abstract", title: "Abstract", subtitle: "Heading, title and abstract text", Component: Abstract },
      { id: "toc", title: "Table of Contents", subtitle: "Contents page, chapters, page numbers", Component: TableOfContents },
    ],
  },
  {
    label: "Figures & tables",
    steps: [
      { id: "captions", title: "Captions", subtitle: "Insert Figure 3.1 / Table 3.1 captions", Component: CaptionCrossRefPanel },
      { id: "lists", title: "List of Figures & Tables", subtitle: "Numbered by chapter, with page numbers", Component: CaptionLists },
    ],
  },
  {
    label: "Report body",
    steps: [{ id: "content", title: "Content", subtitle: "Fonts, 1.5 spacing and justify", Component: Content }],
  },
];

const ALL_IDS = GROUPS.flatMap((g) => g.steps.map((s) => s.id));

export default function App() {
  const styles = useStyles();
  const [openItems, setOpenItems] = React.useState(["front-cover"]);
  const [showNews, setShowNews] = React.useState(false);

  const allOpen = openItems.length === ALL_IDS.length;
  let number = 0;

  const renderItem = (step, numbered) => {
    if (numbered) number += 1;
    const isOpen = openItems.includes(step.id);
    const { Component } = step;
    return (
      <AccordionItem
        key={step.id}
        value={step.id}
        className={mergeClasses(styles.item, isOpen && styles.itemOpen)}
      >
        <AccordionHeader expandIconPosition="end" size="large">
          <div className={styles.headerRow}>
            <span className={mergeClasses(styles.badge, !numbered && styles.badgeInfo)}>
              {numbered ? number : "i"}
            </span>
            <span className={styles.titles}>
              <span className={styles.title}>{step.title}</span>
              <span className={styles.subtitle}>{step.subtitle}</span>
            </span>
          </div>
        </AccordionHeader>
        <AccordionPanel className={styles.panel}>
          <Component />
        </AccordionPanel>
      </AccordionItem>
    );
  };

  return (
    <div className={styles.root}>
      <Header />

      <div className={styles.intro}>
        <Text className={styles.introText}>Work through the steps in order, starting from 1.</Text>
        <div style={{ display: "flex", flexShrink: 0 }}>
          <Button
            size="small"
            appearance="subtle"
            style={{ whiteSpace: "nowrap" }}
            onClick={() => setShowNews(true)}>
            What&apos;s new
          </Button>
          <Button
            size="small"
            appearance="subtle"
            style={{ whiteSpace: "nowrap" }}
            onClick={() => setOpenItems(allOpen ? [] : ALL_IDS)}>
            {allOpen ? "Collapse all" : "Expand all"}
          </Button>
        </div>
      </div>

      <WhatsNew forceOpen={showNews} onClose={() => setShowNews(false)} />

      <Accordion
        multiple
        collapsible
        openItems={openItems}
        onToggle={(_, data) => setOpenItems(data.openItems)}
      >
        {GROUPS.map((group) => (
          <div key={group.label}>
            <Text className={styles.groupLabel}>{group.label}</Text>
            <div className={styles.list}>{group.steps.map((s) => renderItem(s, true))}</div>
          </div>
        ))}
      </Accordion>

      <Text className={styles.footer}>eFormat FYP Report · Universiti Poly-Tech Malaysia</Text>
    </div>
  );
}
