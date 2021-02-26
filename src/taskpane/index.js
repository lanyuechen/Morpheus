import "office-ui-fabric-react/dist/css/fabric.min.css";
import App from "./Entry";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";
/* global AppCpntainer, Component, document, Office, module, React, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "Contoso Task Pane Add-in";

const render = Component => {
  ReactDOM.render(
    <AppContainer>
      <Component title={title} isOfficeInitialized={isOfficeInitialized} />
    </AppContainer>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.initialize = () => {
  isOfficeInitialized = true;
  render(App);
};

/* Initial render showing a progress bar */
render(App);

if (module.hot) {
  module.hot.accept("./Entry", () => {
    const NextApp = require("./Entry").default;
    render(NextApp);
  });
}
