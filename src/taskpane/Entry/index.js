import React from "react";
import { PrimaryButton } from "@fluentui/react";
import Progress from "../components/Progress";

const App = (props) => {
  const { title, isOfficeInitialized } = props;

  const helloWorld = async () => {
    return Word.run(async context => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraph.font.color = "red";

      await context.sync();
    });
  };

  if (!isOfficeInitialized) {
    return (
      <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
    );
  }

  return (
    <div className="ms-welcome">
      {/* <h2>墨菲斯的工具集</h2> */}
      <PrimaryButton
        onClick={helloWorld}
      >
        hello world
      </PrimaryButton>
    </div>
  );
}

export default App;
