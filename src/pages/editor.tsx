import * as React from "react";

const EditorPage: React.FC = () => {
  React.useEffect(() => {
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
      const message = JSON.parse(arg.message);

      if (message.type === "response-config") {
        console.log(message);
      }
    }, () => {});

    Office.context.ui.messageParent(JSON.stringify({
      type: "request-config"
    }));
  }, []);



  return <div>Editor Page</div>;
};

export default EditorPage;
