import { createRoot } from "react-dom/client";
import { registerActions } from "../commands/actions";
import { App } from "./App";
import { pane } from "./store";
import { tabFromUrl, viewFromUrl } from "./views";

function start() {
  pane.init(viewFromUrl(location.search));
  // Ribbon buttons run in this page (shared runtime), so their functions are registered here.
  if (typeof Office !== "undefined" && Office.actions) registerActions();
  createRoot(document.getElementById("root")!).render(<App tab={tabFromUrl(location.search)} />);
}

// Inside PowerPoint, wait for Office.js; opened in a plain browser, render anyway so the
// page (and Diagnostics) still shows something useful.
if (typeof Office !== "undefined") {
  Office.onReady(start);
} else {
  start();
}
