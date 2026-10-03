import { createRoot } from "react-dom/client";
import { App } from "./App";
import { tabFromUrl, viewFromUrl } from "./views";

function start() {
  createRoot(document.getElementById("root")!).render(<App initialView={viewFromUrl(location.search)} tab={tabFromUrl(location.search)} />);
}

// Inside PowerPoint, wait for Office.js; opened in a plain browser, render anyway so the
// page (and Diagnostics) still shows something useful.
if (typeof Office !== "undefined") {
  Office.onReady(start);
} else {
  start();
}
