import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webDarkTheme, webLightTheme } from "@fluentui/react-components";

/* global document, window, Office, module, require */

const title = "UPTM REPORT FORMATTING";

// "ResizeObserver loop completed with undelivered notifications" is a
// harmless browser warning (Fluent UI fields/accordions resizing inside the
// narrow task pane). Nothing breaks, but the webpack dev server shows it as a
// full-screen "Uncaught runtime errors" overlay. Stop it here, before the
// overlay sees it. (Capture-phase listeners on window run first.)
const IGNORED_ERRORS = /ResizeObserver loop (completed with undelivered notifications|limit exceeded)/;
window.addEventListener(
  "error",
  (event) => {
    if (IGNORED_ERRORS.test(event?.message || "")) {
      event.stopImmediatePropagation();
      event.preventDefault();
    }
  },
  true
);

const rootElement = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  root?.render(
    <FluentProvider theme={webDarkTheme}>
      <App title={title} />
    </FluentProvider>
  );
});

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}
