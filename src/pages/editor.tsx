import { makeStyles } from "@fluentui/react-components";
import { DocumentEditor } from "@onlyoffice/document-editor-react";
import * as React from "react";

const base64ToUint8Array = (base64: string) => {
  const binary = atob(base64);

  return new Uint8Array(Array.from(binary, (char) => char.charCodeAt(0)));
}

const useStyles = makeStyles({
  root: {
    height: "100%",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    textAlign: "center",
  }
});

const EditorPage: React.FC = () => {
  const styles = useStyles();
  const [documentServerUrl, setDocumentServerUrl] = React.useState();
  const [config, setConfig] = React.useState();

  React.useEffect(() => {
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
      const message = JSON.parse(arg.message);

      if (message.type === "response-config") {
        const thisDocumentServerUrl = message.data.documentServerUrl;
        const thisConfig = message.data.config;
        const content = message.data.content;
        
        thisConfig.events = {
          onAppReady: () => {
            const editor = window.DocEditor?.instances["onlyofficeEditor"];

            if (editor) {
              const body = base64ToUint8Array(content);
              //@ts-expect-error
              editor.openDocument(body);
            }
          }
        }

        setDocumentServerUrl(thisDocumentServerUrl);
        setConfig(thisConfig);
      }
    }, () => {});

    Office.context.ui.messageParent(JSON.stringify({
      type: "request-config"
    }));
  }, []);

  return (
    <div className={styles.root}>
      {documentServerUrl && config &&
        <DocumentEditor
            id="onlyofficeEditor"
            documentServerUrl={documentServerUrl}
            config={config}
            height="100%"
            width="100%"
          />
        }
    </div>
  ); 
};

export default EditorPage;
