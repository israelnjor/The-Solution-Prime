import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import { isDesktop } from "./platform";
import "prosemirror-view/style/prosemirror.css";
import "./styles.css";

// In the desktop app, a right-click outside the document would otherwise show the web view's own menu
// (Back, Refresh, Share, Save as…), which belongs to a browser, not to Prime Word. Text boxes keep
// theirs (Cut, Copy, Paste); the document has its own menu.
if (isDesktop) {
  window.addEventListener("contextmenu", (e) => {
    const t = e.target as HTMLElement | null;
    if (t?.closest("input, textarea")) return;
    e.preventDefault();
  });
}

createRoot(document.getElementById("root")!).render(<StrictMode><App /></StrictMode>);
