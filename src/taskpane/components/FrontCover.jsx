import React from "react";
import {
  makeStyles,
  useId,
  Dropdown,
  Option,
  Button,
  TableBody,
  TableCell,
  TableRow,
  Table,
  TableHeader,
  TableHeaderCell,
  TableCellLayout,
  PresenceBadgeStatus,
  Avatar,
} from "@fluentui/react-components";

const useStyles = makeStyles({
  MostFrontCoverWrapper: {
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

  label: {
    marginBottom: "10px",
  },

  contentMostFrontCover: {
    paddingLeft: "25px",
  },

  button: {
    marginTop: "10px",
  },

  wrapButton: {
    display: "flex",
    justifyContent: "flex-end",
    marginRight: "40px",
  },
});

export default function FrontCover() {
  const styles = useStyles();

  // table for margin
  const tableMargin = [
    { columnKey: "element", label: "Element" },
    { columnKey: "formatStyle", label: "Format Style" },
  ];

  // table for element
  const tableElement = [
    { columnKey: "element", label: "Element" },
    { columnKey: "formatStyle", label: "Format Style" },
  ];

  // items for margin format
  const itemsMargin = [
    {
      margin: { element: "Top Margin", format: "4 cm" },
    },
    {
      margin: { element: "Bottom Margin", format: "4 cm" },
    },
    {
      margin: { element: "Left Margin", format: "3.2 cm" },
    },
    {
      margin: { element: "Right Margin", format: "3.2 cm" },
    },
  ];

  // items for element format
  const itemsElement = [
    {
      element: {
        label: "Title of Project",
        format: {
          fontStyle: "Times New Roman 18",
          line: "60mm from the top",
        },
      },
    },
    {
      element: {
        label: "Student's Full Name",
        format: {
          fontStyle: "Times New Roman 18",
          line: "Center of Page",
        },
      },
    },
    {
      element: {
        label: "UNIVERSITI POLY-TECH MALAYSIA",
        format: {
          fontStyle: "Times New Roman 18",
          line: "60mm from the bottom",
        },
      },
    },
  ];

  return (
    <>
      <div className={styles.MostFrontCoverWrapper}>
        <h2 className={styles.heading}>Front Cover</h2>

        {/* table margin*/}
        <div className={styles.contentMostFrontCover}>
          <p className={styles.label}>Guide for Margin</p>
          <Table arial-label="Default table" style={{ maxWidth: "200px" }}>
            <TableHeader>
              <TableRow>
                {tableMargin.map((column) => (
                  <TableHeaderCell key={column.columnKey}>{column.label}</TableHeaderCell>
                ))}
              </TableRow>
            </TableHeader>
            <TableBody>
              {itemsMargin.map((item) => (
                <TableRow key={item.margin.element}>
                  <TableCell>
                    <TableCellLayout>{item.margin.element}</TableCellLayout>
                  </TableCell>

                  <TableCell>{item.margin.format}</TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>

          {/* table element*/}
          <p className={styles.label}>Guide for Element</p>
          <Table arial-label="Default table" style={{ maxWidth: "250px" }}>
            <TableHeader>
              <TableRow>
                {tableElement.map((column) => (
                  <TableHeaderCell key={column.columnKey}>{column.label}</TableHeaderCell>
                ))}
              </TableRow>
            </TableHeader>
            <TableBody>
              {itemsElement.map((item) => (
                <TableRow key={item.element.label}>
                  <TableCell>
                    <TableCellLayout>{item.element.label}</TableCellLayout>
                  </TableCell>

                  <TableCell>
                    <TableCellLayout description={item.element.format.line}>
                      {item.element.format.fontStyle}
                    </TableCellLayout>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </div>
      </div>
    </>
  );
}
