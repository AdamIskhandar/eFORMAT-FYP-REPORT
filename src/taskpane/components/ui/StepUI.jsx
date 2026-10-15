import React from "react";
import { Dismiss16Regular } from "@fluentui/react-icons";
import {
  makeStyles,
  tokens,
  Text,
  MessageBar,
  MessageBarBody,
  MessageBarActions,
  Button,
  Table,
  TableHeader,
  TableHeaderCell,
  TableBody,
  TableRow,
  TableCell,
  ProgressBar,
  Spinner,
} from "@fluentui/react-components";

/**
 * Shared building blocks for every step panel, so all steps look and behave
 * the same: a column body, small grey hints, sub headings, a button row,
 * a status box and a guide table. UI only - no Word logic in here.
 */

const useStyles = makeStyles({
  body: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    padding: "4px 4px 12px 4px",
    minWidth: 0,
    // Fluent's Dropdown has a fixed 250px min-width, wider than the task
    // pane - make every field fill (and never overflow) its column instead.
    "& .fui-Dropdown, & .fui-Input, & .fui-Combobox, & .fui-Field": {
      minWidth: 0,
      width: "100%",
      maxWidth: "100%",
      boxSizing: "border-box",
    },
  },
  hint: {
    display: "block",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
  },
  sub: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    margin: "6px 0 0 0",
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  subLine: {
    flex: 1,
    height: "1px",
    backgroundColor: tokens.colorNeutralStroke2,
  },
  actions: {
    display: "flex",
    gap: "8px",
    flexWrap: "wrap",
    justifyContent: "flex-end",
  },
  actionsStretch: {
    display: "grid",
    gridTemplateColumns: "1fr",
    gap: "8px",
  },
  group: {
    display: "flex",
    flexDirection: "column",
    gap: "10px",
    minWidth: 0,
    padding: "12px",
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  tableTitle: {
    display: "block",
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    marginBottom: "4px",
  },
  table: {
    width: "100%",
  },
  busy: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    padding: "10px 12px",
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorBrandBackground2,
    border: `1px solid ${tokens.colorBrandStroke2}`,
  },
  busyRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  busyHint: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground2,
  },
  cellFormat: {
    color: tokens.colorNeutralForeground2,
  },
});

export function StepBody({ children }) {
  const styles = useStyles();
  return <div className={styles.body}>{children}</div>;
}

export function Hint({ children }) {
  const styles = useStyles();
  return <Text className={styles.hint}>{children}</Text>;
}

export function SubHeading({ children }) {
  const styles = useStyles();
  return (
    <div className={styles.sub}>
      <span>{children}</span>
      <span className={styles.subLine} />
    </div>
  );
}

// A boxed group of controls that belong together (e.g. Header / Footer).
export function Group({ children }) {
  const styles = useStyles();
  return <div className={styles.group}>{children}</div>;
}

// Buttons, right-aligned; `stretch` makes them full width (main actions).
export function Actions({ children, stretch = false }) {
  const styles = useStyles();
  return <div className={stretch ? styles.actionsStretch : styles.actions}>{children}</div>;
}

// tone: "success" | "error" | "warning" | "info". Pass onDismiss to show a close button.
export function Status({ text, tone = "info", onDismiss }) {
  if (!text) return null;
  return (
    <MessageBar intent={tone} layout="multiline">
      <MessageBarBody>{text}</MessageBarBody>
      {onDismiss && (
        <MessageBarActions
          containerAction={
            <Button
              appearance="transparent"
              size="small"
              icon={<Dismiss16Regular />}
              aria-label="Dismiss"
              onClick={onDismiss}
            />
          }
        />
      )}
    </MessageBar>
  );
}

export function Notice({ children, tone = "warning" }) {
  return (
    <MessageBar intent={tone} layout="multiline">
      <MessageBarBody>{children}</MessageBarBody>
    </MessageBar>
  );
}

// rows: [{ element, format }]
export function GuideTable({ title, rows }) {
  const styles = useStyles();
  return (
    <div>
      {title && <Text className={styles.tableTitle}>{title}</Text>}
      <Table size="extra-small" aria-label={title || "Format guide"} className={styles.table}>
        <TableHeader>
          <TableRow>
            <TableHeaderCell>Element</TableHeaderCell>
            <TableHeaderCell>Format</TableHeaderCell>
          </TableRow>
        </TableHeader>
        <TableBody>
          {rows.map((r) => (
            <TableRow key={r.element}>
              <TableCell>{r.element}</TableCell>
              <TableCell className={styles.cellFormat}>{r.format}</TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </div>
  );
}

/**
 * Loading state for slow Word actions.
 *   const busy = useBusy();
 *   <Button disabled={busy.active} onClick={busy.wrap("Generating...", handleGenerate, { settle: 4200 })}>
 *   <BusyBar busy={busy} />
 * wrap() runs the handler unchanged; it only shows the bar while it runs.
 * `settle` keeps a second "finishing" phase on screen for that many ms
 * after the handler returns (e.g. while page numbers refresh).
 */
export function useBusy() {
  const [state, setState] = React.useState(null); // { label, phase: "working" | "settling" }
  const timer = React.useRef(null);
  React.useEffect(() => () => clearTimeout(timer.current), []);

  const wrap =
    (label, handler, { settle = 0, settleLabel = "Updating page numbers...", group = null } = {}) =>
    async (...args) => {
      clearTimeout(timer.current);
      setState({ label, phase: "working", group });
      try {
        await handler(...args);
      } finally {
        if (settle > 0) {
          setState({ label: settleLabel, phase: "settling", group });
          timer.current = setTimeout(() => setState(null), settle);
        } else {
          setState(null);
        }
      }
    };

  return { active: !!state && state.phase === "working", state, wrap };
}

// `group`: only show for actions wrapped with the same group (so the bar
// appears next to the button that was clicked).
export function BusyBar({ busy, group = null }) {
  const styles = useStyles();
  if (!busy.state || (group && busy.state.group !== group)) return null;
  const working = busy.state.phase === "working";
  return (
    <div className={styles.busy} role="status" aria-live="polite">
      <div className={styles.busyRow}>
        <Spinner size="extra-tiny" />
        <span>{busy.state.label}</span>
      </div>
      <ProgressBar thickness="large" />
      <span className={styles.busyHint}>
        {working
          ? "Please wait and don't type in Word until this finishes. Big reports can take a while."
          : "Almost done - you can keep working."}
      </span>
    </div>
  );
}
