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
} from "@fluentui/react-components";

const useStyles = makeStyles({
  SideCoverWrapper: {
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

  contentSideCover: {
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

export default function SideCover() {
  const styles = useStyles();

  // table for element
  const tableElement = [
    { columnKey: "element", label: "Element" },
    { columnKey: "formatStyle", label: "Format Style" },
  ];

  // items for element format
  const itemsElement = [
    {
      element: {
        label: "Student ID",
        format: {
          fontStyle: "ARIAL 11",
        },
      },
    },
    {
      element: {
        label: "Category",
        format: {
          fontStyle: "ARIAL 11",
        },
      },
    },
    {
      element: {
        label: "Year",
        format: {
          fontStyle: "ARIAL 11",
        },
      },
    },
  ];

  return (
    <>
      <div className={styles.SideCoverWrapper}>
        <h2 className={styles.heading}>Side Cover</h2>

        <div className={styles.contentSideCover}>
          {/* table element*/}
          <p className={styles.label}>Guide for Side Covert</p>
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
                    <TableCellLayout>{item.element.format.fontStyle}</TableCellLayout>
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
