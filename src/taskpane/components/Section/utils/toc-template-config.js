/**
 * Report Structure Config (hardcoded, admin-edited)
 *
 * Single source of truth for the Table of Contents page AND the real
 * headings created throughout the document. To change the required report
 * structure, edit this list and redeploy - no database or admin UI needed.
 *
 * level 1 = chapter / Appendix / References  (Heading 1: centered, Arial 18)
 * level 2 = section, e.g. 1.1                (Heading 2: left,     Arial 14)
 * level 3 = subsection, e.g. 1.3.1           (Heading 3: indented, Arial 12)
 * level 4 = sub-subsection, e.g. 3.5.1.1     (Heading 4: indented, Arial 12 italic)
 *
 * The heading text must start with its number ("3.3 Phases ...") - the
 * List of Figures/Tables numbering reads the number from the heading text.
 *
 * placeholder: optional hint line shown under the heading. Only give it to
 * headings the student actually writes under (those without sub-headings).
 */

const WRITE = "Write your content here.";
const entry = (level, text, placeholder) =>
  placeholder ? { level, text, placeholder } : { level, text };

export const REPORT_STRUCTURE = [
  // ---- 1 INTRODUCTION ----
  entry(1, "1 INTRODUCTION"),
  entry(2, "1.1 Introduction", WRITE),
  entry(2, "1.2 Project Background", WRITE),
  entry(2, "1.3 Problem Statement"),
  entry(3, "1.3.1 UPTM use a manual way to record asset data (example)", WRITE),
  entry(3, "1.3.2 Problem statement no 2", WRITE),
  entry(3, "1.3.3 Third Problem Statement", WRITE),
  entry(2, "1.4 Project Objectives"),
  entry(3, "1.4.1 To Investigate possible solutions, etc, etc (example)", WRITE),
  entry(3, "1.4.2 To Analyse and model the requirements (example)", WRITE),
  entry(3, "1.4.3 To develop an asset management system for UPTM (example)", WRITE),
  entry(2, "1.5 Scope and Target User"),
  entry(3, "1.5.1 Project Scope", WRITE),
  entry(3, "1.5.2 Product Scope", WRITE),
  entry(3, "1.5.3 Target User", WRITE),
  entry(2, "1.6 Conclusion", WRITE),

  // ---- 2 LITERATURE REVIEW ----
  entry(1, "2 LITERATURE REVIEW"),
  entry(2, "2.1 Introduction", WRITE),
  entry(2, "2.2 Investigation"),
  entry(3, "2.2.1 Definition SQL Injection and Prevention (EXAMPLE)", WRITE),
  entry(3, "2.2.2 SQL Injection (EXAMPLE)", WRITE),
  entry(3, "2.2.3 Type of SQL Injection (EXAMPLE)", WRITE),
  entry(2, "2.3 Related Works"),
  entry(3, "2.3.1 Project A", WRITE),
  entry(3, "2.3.2 Project B", WRITE),
  entry(3, "2.3.3 Project C", WRITE),
  entry(2, "2.4 Comparison", WRITE),
  entry(2, "2.5 Discussion", WRITE),
  entry(2, "2.6 Conclusion", WRITE),

  // ---- 3 METHODOLOGY ----
  entry(1, "3 METHODOLOGY"),
  entry(2, "3.1 Introduction", WRITE),
  entry(2, "3.2 Agile Methodology", WRITE),
  entry(2, "3.3 Phases in Agile Methodology"),
  entry(3, "3.3.1 Phase 1", WRITE),
  entry(3, "3.3.2 Phase 2", WRITE),
  entry(3, "3.3.3 Phase 3", WRITE),
  entry(3, "3.3.4 Phase 4", WRITE),
  entry(2, "3.4 Requirement"),
  entry(3, "3.4.1 Data Gathering Techniques", WRITE),
  entry(3, "3.4.2 Functional Requirement", WRITE),
  entry(3, "3.4.3 Non-Function Requirement", WRITE),
  entry(3, "3.4.4 System Requirement", WRITE),
  entry(2, "3.5 Analysis"),
  entry(3, "3.5.1 Data Gathering Analysis"),
  entry(4, "3.5.1.1 Questionnaire Analysis", WRITE),
  entry(4, "3.5.1.2 Interview Analysis", WRITE),
  entry(3, "3.5.2 Use Case Model", WRITE),
  entry(3, "3.5.3 Flowchart", WRITE),
  entry(3, "3.5.4 BPNM (Business Process Modelling Notation)", WRITE),
  entry(2, "3.6 Conclusion", WRITE),

  // ---- 4 DESIGN ----
  entry(1, "4 DESIGN"),
  entry(2, "4.1 Introduction", WRITE),
  entry(2, "4.2 Interface Design - Wireframe", WRITE),
  entry(2, "4.3 Database Design"),
  entry(3, "4.3.1 Data Dictionary", WRITE),
  entry(3, "4.3.2 Data Flow Diagram (DFD)", WRITE),
  entry(3, "4.3.3 Entity Relational Diagram (ERD)", WRITE),
  entry(2, "4.4 Security System Framework (CT206)", WRITE),
  entry(2, "4.5 Implementation"),
  entry(3, "4.5.1 Implementation Tools", WRITE),
  entry(3, "4.5.2 System Interface", WRITE),
  entry(3, "4.5.3 Significant Functions (i.e. security element)", WRITE),
  entry(2, "4.6 Conclusion", WRITE),

  // ---- 5 FINDING ----
  entry(1, "5 FINDING"),
  entry(2, "5.1 Introduction", WRITE),
  entry(2, "5.2 Non-Functional Testing"),
  entry(3, "5.2.1 Performance Testing", WRITE),
  entry(3, "5.2.2 Security Testing", WRITE),
  entry(3, "5.2.3 Usability Testing", WRITE),
  entry(2, "5.3 Functional Testing"),
  entry(3, "5.3.1 Unit Testing", WRITE),
  entry(3, "5.3.2 Integration Testing", WRITE),
  entry(3, "5.3.3 System Testing", WRITE),
  entry(2, "5.4 Acceptance Testing"),
  entry(3, "5.4.1 Client Acceptance Testing", WRITE),
  entry(3, "5.4.2 User Acceptance Testing", WRITE),
  entry(2, "5.5 Conclusion", WRITE),

  // ---- 6 CONCLUSION ----
  entry(1, "6 CONCLUSION"),
  entry(2, "6.1 Introduction", WRITE),
  entry(2, "6.2 Project Schedule"),
  entry(3, "6.2.1 Work Breakdown Structure", WRITE),
  entry(3, "6.2.2 Gantt Chart", WRITE),
  entry(2, "6.3 Risk Management", WRITE),
  entry(2, "6.4 Achievement"),
  entry(3, "6.4.1 First Point of the Objective", WRITE),
  entry(3, "6.4.2 Second Point of the Objective", WRITE),
  entry(3, "6.4.3 Third Point of the Objective", WRITE),
  entry(2, "6.5 Constraint and Limitation", WRITE),
  entry(2, "6.6 Future Work and Recommendation", WRITE),
  entry(2, "6.7 Conclusion", WRITE),

  // ---- Back matter ----
  entry(
    1,
    "Appendix A \u2013 Requirements Specification Document",
    "Insert appendix content here."
  ),
  entry(1, "Appendix B \u2013 User Manual", "Insert appendix content here."),
  entry(1, "Appendix C \u2013 Turnitin Result", "Insert appendix content here."),
  entry(1, "Appendix D \u2013 Log Book", "Insert appendix content here."),
  entry(1, "References", "Insert your reference list here."),
];
