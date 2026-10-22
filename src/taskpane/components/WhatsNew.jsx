import * as React from "react";
import {
  makeStyles,
  tokens,
  Button,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  MessageBarActions,
} from "@fluentui/react-components";
import { Dismiss16Regular } from "@fluentui/react-icons";
import { UPDATES, LATEST } from "../whats-new";

/* global window */

// The last version this student has seen, kept in this browser/Word only.
const SEEN_KEY = "eformat-whats-new-seen";

function readSeen() {
  try {
    return window.localStorage.getItem(SEEN_KEY);
  } catch (e) {
    return null;
  }
}

function writeSeen(version) {
  try {
    window.localStorage.setItem(SEEN_KEY, version);
  } catch (e) {
    // Storage blocked - the banner just shows again next time.
  }
}

const useStyles = makeStyles({
  wrap: {
    padding: "0 10px 8px 10px",
  },
  list: {
    margin: "6px 0 0 0",
    paddingLeft: "18px",
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
  },
  older: {
    marginTop: "8px",
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
  },
  date: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  // The banner's date: on its own line, under the title.
  dateLine: {
    display: "block",
    marginTop: "2px",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightRegular,
  },
  title: {
    display: "block",
  },
});

/**
 * "What's new" banner. Shows by itself once per new version (see
 * whats-new.js); `forceOpen` lets the What's new button reopen it.
 */
export default function WhatsNew({ forceOpen = false, onClose }) {
  const styles = useStyles();
  const [unseen, setUnseen] = React.useState(() => readSeen() !== LATEST.version);
  const [expanded, setExpanded] = React.useState(false);

  React.useEffect(() => {
    if (forceOpen) setExpanded(true);
  }, [forceOpen]);

  if (!unseen && !forceOpen) return null;

  const close = () => {
    writeSeen(LATEST.version);
    setUnseen(false);
    setExpanded(false);
    if (onClose) onClose();
  };

  return (
    <div className={styles.wrap}>
      <MessageBar intent="info" layout="multiline">
        <MessageBarBody>
          <MessageBarTitle className={styles.title}>What&apos;s new in eFormat {LATEST.version}</MessageBarTitle>
          <span className={styles.dateLine}>Updated {LATEST.date}</span>
          {expanded && (
            <>
              <ul className={styles.list}>
                {LATEST.changes.map((c) => (
                  <li key={c}>{c}</li>
                ))}
              </ul>
              {UPDATES.slice(1).map((u) => (
                <div key={u.version}>
                  <div className={styles.older}>
                    {u.version} <span className={styles.date}>- {u.date}</span>
                  </div>
                  <ul className={styles.list}>
                    {u.changes.map((c) => (
                      <li key={c}>{c}</li>
                    ))}
                  </ul>
                </div>
              ))}
            </>
          )}
        </MessageBarBody>
        <MessageBarActions
          containerAction={
            <Button
              appearance="transparent"
              size="small"
              icon={<Dismiss16Regular />}
              aria-label="Close"
              onClick={close}
            />
          }
        >
          <Button size="small" onClick={() => setExpanded((v) => !v)}>
            {expanded ? "Hide changes" : "See changes"}
          </Button>
          <Button size="small" appearance="primary" onClick={close}>
            Got it
          </Button>
        </MessageBarActions>
      </MessageBar>
    </div>
  );
}
