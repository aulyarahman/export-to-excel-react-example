import React from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import "./index.css";
import { PagesExport } from "./exports-data";

ReactDOM.createRoot(document.getElementById("root") as HTMLElement).render(
  <React.StrictMode>
    {/* <App /> */}
    <PagesExport />
  </React.StrictMode>
);
