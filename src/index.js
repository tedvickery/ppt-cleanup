import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";

function renderApp() {
  const root = ReactDOM.createRoot(document.getElementById("root"));
  root.render(<React.StrictMode><App /></React.StrictMode>);
}

// Office.onReady is the modern recommended way — works in Office Online and Desktop
if (window.Office) {
  window.Office.onReady(() => renderApp());
} else {
  // Plain browser fallback for testing
  renderApp();
}
