import React, { useState, useEffect, useCallback } from "react";
import {
  Button,
  Badge,
  ProgressBar,
  Card,
  CardHeader,
  Text,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  MessageBarActions,
  Divider,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  getComplianceStatus,
  getFormattingProgress,
  getNotifications,
  dismissNotification,
  clearAllNotifications,
  refreshComplianceNotifications,
} from "../Section/utils/dashboard-utils";

const useStyles = makeStyles({
  panel: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingHorizontalM,
  },
  sectionTitle: {
    marginTop: tokens.spacingVerticalL,
    marginBottom: 0,
  },
  buttonRow: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    flexWrap: "wrap",
  },
  scoreRow: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalM,
    marginBottom: "5px",
    marginTop: "5px",
  },
  areaCard: {
    display: "flex",
    flexDirection: "row",
    alignItems: "center",
    justifyContent: "space-between",
    padding: tokens.spacingHorizontalS,
  },
  progressStep: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
  },
  progressStepRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
  },
  notificationList: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXS,
    marginTop: "10px",
  },
  ContentWrapper: {
    paddingLeft: "10px",
  },
  contentCover: {
    paddingLeft: "25px",
  },
  heading: {
    color: "white",
    fontSize: "20px",
  },
  label: {
    marginBottom: "10px",
  },
  labelCompliance: {
    marginTop: "20px",
  },
  labelNotification: {
    marginTop: "20px",
    marginBottom: "100px",
  },
  wrapButtonSize: {
    display: "flex",
    justifyContent: "flex-end",
    marginRight: "40px",
    maxWidth: "250px",
    minWidth: "250px",
  },
  messageBar: {
    marginTop: "10px",
  },
});

/**
 * Task 7.1 - User Dashboard Development.
 * Shows compliance status, formatting progress, and notifications for
 * the current document. Slots in as section "13)".
 */
export default function DashboardPanel() {
  const styles = useStyles();

  const [compliance, setCompliance] = useState(null);
  const [progress, setProgress] = useState(null);
  const [notifications, setNotifications] = useState([]);
  const [status, setStatus] = useState(null);

  const runInWord = useCallback(async (fn) => {
    try {
      await Word.run(fn);
    } catch (err) {
      console.error(err);
      setStatus({ intent: "error", text: err.message });
    }
  }, []);

  const loadDashboard = useCallback(() => {
    runInWord(async (context) => {
      // Sequential, not Promise.all: each of these does its own
      // load()/sync() cycle on the shared context, and running them
      // concurrently causes intermittent "call load before reading"
      // errors from the Word API.
      const complianceResult = await getComplianceStatus(context);
      const progressResult = await getFormattingProgress(context);
      const notificationsResult = await getNotifications(context);
      setCompliance(complianceResult);
      setProgress(progressResult);
      setNotifications(notificationsResult);
    });
  }, [runInWord]);

  useEffect(() => {
    loadDashboard();
  }, [loadDashboard]);

  const handleRefresh = () => {
    runInWord(async (context) => {
      await refreshComplianceNotifications(context);
    }).then(loadDashboard);
  };

  const handleDismiss = (id) => {
    runInWord(async (context) => {
      const remaining = await dismissNotification(context, id);
      setNotifications(remaining);
    });
  };

  const handleClearAll = () => {
    runInWord(async (context) => {
      await clearAllNotifications(context);
      setNotifications([]);
    });
  };

  return (
    <div className={styles.ContentWrapper}>
      <h3 className={styles.heading}>17) Dashboard</h3>

      <div className={styles.contentCover}>
        <div className={styles.buttonRow}>
          <Button appearance="secondary" onClick={loadDashboard}>
            Refresh Dashboard
          </Button>
          <Button appearance="secondary" onClick={handleRefresh}>
            Check Compliance Now
          </Button>
        </div>

        <div className={styles.labelCompliance}>
          {/* Task 7.1 - Display document compliance status */}
          <div className={styles.progressStep}>
            <Text weight="semibold">Compliance Status</Text>
            {compliance && (
              <>
                <div className={styles.scoreRow}>
                  <ProgressBar value={compliance.overallScore / 100} />
                  <Text>{compliance.overallScore}%</Text>
                </div>
                {compliance.areas.map((area) => (
                  <Card key={area.key} className={styles.areaCard}>
                    <Text>{area.label}</Text>
                    <Badge
                      appearance="filled"
                      color={
                        area.status === "pass"
                          ? "success"
                          : area.status === "empty"
                            ? "informative"
                            : "danger"
                      }
                    >
                      {area.status === "pass"
                        ? "Pass"
                        : area.status === "empty"
                          ? "Not started"
                          : `${area.issues.length} issue(s)`}
                    </Badge>
                  </Card>
                ))}
              </>
            )}
          </div>
        </div>

        <div className={styles.labelCompliance}>
          {/* Task 7.1 - Display formatting progress indicators */}
          <div className={styles.progressStep}>
            <Text weight="semibold">Formatting Progress</Text>
            {progress && (
              <>
                <div className={styles.scoreRow}>
                  <ProgressBar value={progress.percentComplete / 100} />
                  <Text>{progress.percentComplete}%</Text>
                </div>
                {progress.steps.map((step) => (
                  <div key={step.key} className={styles.progressStepRow}>
                    <Text>{step.label}</Text>
                    <Badge appearance="tint" color={step.done ? "success" : "subtle"}>
                      {step.done ? "Done" : "Pending"}
                    </Badge>
                  </div>
                ))}
              </>
            )}
          </div>
        </div>

        <div className={styles.labelNotification}>
          {/* Task 7.1 - Develop notification system */}
          <div className={styles.progressStepRow}>
            <Text weight="semibold">Notifications</Text>
            {notifications.length > 0 && (
              <Button size="small" appearance="transparent" onClick={handleClearAll}>
                Clear All
              </Button>
            )}
          </div>

          <div className={styles.notificationList}>
            {notifications.length === 0 && <Text size={200}>No notifications.</Text>}
            {notifications.map((n) => (
              <MessageBar key={n.id} intent={n.severity === "warning" ? "warning" : n.severity}>
                <MessageBarBody>
                  <MessageBarTitle>{new Date(n.timestamp).toLocaleString()}</MessageBarTitle>
                  {n.message}
                </MessageBarBody>
                <MessageBarActions>
                  <Button size="small" appearance="transparent" onClick={() => handleDismiss(n.id)}>
                    Dismiss
                  </Button>
                </MessageBarActions>
              </MessageBar>
            ))}
          </div>

          {status && (
            <MessageBar intent="error">
              <MessageBarBody>
                <MessageBarTitle>Error</MessageBarTitle>
                {status.text}
              </MessageBarBody>
            </MessageBar>
          )}
        </div>
      </div>
    </div>
  );
}
