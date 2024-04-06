import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";

const rootElement = document.getElementById("container");
const root = createRoot(rootElement);

Office.onReady(() => {
  root.render(<App />);
});

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root.render(<NextApp />);
  });
}
