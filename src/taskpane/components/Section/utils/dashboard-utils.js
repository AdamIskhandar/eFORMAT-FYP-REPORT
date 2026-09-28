/**
 * User Dashboard Module (Task 7.1)
 * - Dashboard interface (data layer — see DashboardPanel.fluent.jsx for UI)
 * - Display document compliance status
 * - Display formatting progress indicators
 * - Notification system
 *
 * Pulls status from the modules built in earlier tasks (captions,
 * citations) and tracks its own lightweight "progress" and
 * "notifications" records in the document's Settings store, the same
 * pattern used by citation-utils.js.
 */

import { validateCaptions } from "./word-caption-utils";
import { validateCitations, getCitationDatabase } from "./citation-utils";

const NOTIFICATIONS_KEY = "notifications";
const PROGRESS_KEY = "formattingProgress";

const PROGRESS_STEPS = [
  { key: "titlePage", label: "Title Page Formatting" },
  { key: "figureCaptions", label: "Figure Captions" },
  { key: "tableCaptions", label: "Table Captions" },
  { key: "citations", label: "In-text Citations" },
  { key: "bibliography", label: "Reference List" },
];

function makeId() {
  return `n${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
}

// ---------------------------------------------------------------------
// Task 7.1 - Display document compliance status
// ---------------------------------------------------------------------

/**
 * Runs the validation routines from earlier modules and returns a single
 * aggregated compliance report the dashboard can render as pass/fail
 * badges per area.
 */
export async function getComplianceStatus(context) {
  // Run sequentially rather than with Promise.all: each of these does its
  // own load()/sync()/read cycle against the same shared context, and
  // interleaving those cycles concurrently causes "call load before
  // reading" errors because one call's sync() can resolve before
  // another's load() has been queued.
  const figureResult = await validateCaptions(context, "figure");
  const tableResult = await validateCaptions(context, "table");
  const citationResult = await validateCitations(context);
  const citationCount = (await getCitationDatabase(context)).length;

  // An empty set trivially satisfies "no inconsistencies found" — that's
  // not the same thing as "this section meets requirements." Distinguish
  // "nothing here yet" from an actual pass so the dashboard doesn't show
  // green on a blank document.
  const counts = {
    figureCaptions: figureResult.count,
    tableCaptions: tableResult.count,
    citations: citationCount,
  };

  function status(key, valid, issues) {
    if (counts[key] === 0) return "empty";
    return valid ? "pass" : "fail";
  }

  const areas = [
    {
      key: "figureCaptions",
      label: "Figure Captions",
      status: status("figureCaptions", figureResult.valid),
      issues: figureResult.issues,
    },
    {
      key: "tableCaptions",
      label: "Table Captions",
      status: status("tableCaptions", tableResult.valid),
      issues: tableResult.issues,
    },
    {
      key: "citations",
      label: "Citations",
      status: status("citations", citationResult.valid),
      issues: citationResult.issues,
    },
  ];

  const passCount = areas.filter((a) => a.status === "pass").length;

  return {
    areas,
    overallScore: Math.round((passCount / areas.length) * 100),
    allValid: passCount === areas.length,
  };
}

// ---------------------------------------------------------------------
// Task 7.1 - Display formatting progress indicators
// ---------------------------------------------------------------------

async function readProgressFlags(context) {
  const setting = context.document.settings.getItemOrNullObject(PROGRESS_KEY);
  setting.load("value");
  await context.sync();
  if (setting.isNullObject) return {};
  try {
    return JSON.parse(setting.value) || {};
  } catch {
    return {};
  }
}

/**
 * Marks a step complete/incomplete. Call this from the relevant module's
 * action handlers (e.g. after generateBibliography succeeds, call
 * markStepComplete(context, "bibliography", true)).
 */
export async function markStepComplete(context, stepKey, complete = true) {
  const flags = await readProgressFlags(context);
  flags[stepKey] = complete;
  context.document.settings.add(PROGRESS_KEY, JSON.stringify(flags));
  await context.sync();
  return flags;
}

/**
 * Computes overall formatting progress. Where possible this is derived
 * directly from document content (captions present, references section
 * present) rather than only trusting stored flags, so the indicator
 * stays accurate even if markStepComplete was never called.
 */
export async function getFormattingProgress(context) {
  const flags = await readProgressFlags(context);

  // Sequential for the same reason as getComplianceStatus above.
  const figureCaptions = await validateCaptions(context, "figure");
  const tableCaptions = await validateCaptions(context, "table");
  const citations = await getCitationDatabase(context);

  const derived = {
    titlePage: !!flags.titlePage,
    figureCaptions: figureCaptions.count > 0,
    tableCaptions: tableCaptions.count > 0,
    citations: citations.length > 0,
    bibliography: !!flags.bibliography,
  };

  const steps = PROGRESS_STEPS.map((step) => ({
    ...step,
    done: derived[step.key],
  }));

  const doneCount = steps.filter((s) => s.done).length;

  return {
    steps,
    percentComplete: Math.round((doneCount / steps.length) * 100),
  };
}

// ---------------------------------------------------------------------
// Task 7.1 - Develop notification system
// ---------------------------------------------------------------------

async function readNotifications(context) {
  const setting = context.document.settings.getItemOrNullObject(NOTIFICATIONS_KEY);
  setting.load("value");
  await context.sync();
  if (setting.isNullObject) return [];
  try {
    return JSON.parse(setting.value) || [];
  } catch {
    return [];
  }
}

async function writeNotifications(context, notifications) {
  context.document.settings.add(NOTIFICATIONS_KEY, JSON.stringify(notifications));
  await context.sync();
}

export async function getNotifications(context) {
  return readNotifications(context);
}

/** severity: "info" | "warning" | "error" | "success" */
export async function addNotification(context, severity, message) {
  const notifications = await readNotifications(context);
  const notification = {
    id: makeId(),
    severity,
    message,
    timestamp: new Date().toISOString(),
    read: false,
  };
  notifications.unshift(notification);
  await writeNotifications(context, notifications.slice(0, 50)); // cap history
  return notification;
}

export async function dismissNotification(context, id) {
  const notifications = await readNotifications(context);
  const remaining = notifications.filter((n) => n.id !== id);
  await writeNotifications(context, remaining);
  return remaining;
}

export async function clearAllNotifications(context) {
  await writeNotifications(context, []);
  return [];
}

/**
 * Convenience: runs compliance + progress checks and raises notifications
 * for anything currently failing, so the notification list doubles as an
 * activity feed of outstanding issues.
 */
export async function refreshComplianceNotifications(context) {
  const compliance = await getComplianceStatus(context);
  const created = [];

  for (const area of compliance.areas) {
    if (area.status === "fail") {
      const n = await addNotification(
        context,
        "warning",
        `${area.label}: ${area.issues[0]}${area.issues.length > 1 ? ` (+${area.issues.length - 1} more)` : ""}`
      );
      created.push(n);
    }
  }

  return created;
}
