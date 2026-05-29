import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";

function renderApp() {
  const root = ReactDOM.createRoot(document.getElementById("root"));
  root.render(<React.StrictMode><App /></React.StrictMode>);
}

// If running inside Office, wait for Office.initialize
// If running in a plain browser (testing), render immediately
if (window.Office && window.Office.initialize !== undefined) {
  window.Office.initialize = () => renderApp();
} else {
  // Fallback: render after a short delay to let Office.js attempt to load
  setTimeout(() => {
    if (document.getElementById("root").childElementCount === 0) {
      renderApp();
    }
  }, 2000);
}
