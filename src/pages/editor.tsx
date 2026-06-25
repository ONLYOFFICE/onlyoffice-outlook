import { makeStyles } from "@fluentui/react-components";
import { DocumentEditor } from "@onlyoffice/document-editor-react";
import * as React from "react";

import { Config } from '@onlyoffice/doceditor-types';

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

const EDITOR_ID = "onlyofficeEditor";

const EditorPage: React.FC = () => {
  const styles = useStyles();
  const [documentServerUrl, setDocumentServerUrl] = React.useState();
  const [config, setConfig] = React.useState<Config | undefined>();
  const attachmentId = React.useRef();

  React.useEffect(() => {
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
      const message = JSON.parse(arg.message);

      if (message.type === "response-config") {
        const thisDocumentServerUrl = message.data.documentServerUrl;
        const thisConfig = message.data.config;
        const thisAttacmentId = message.data.attachmentId;
        const content = message.data.content;
        
        thisConfig.events = {
          onAppReady: () => {
            const editor = window.DocEditor?.instances[EDITOR_ID];

            if (editor) {
              const body = base64ToUint8Array(content);
              //@ts-expect-error
              editor.openDocument(body);
            }
          },
          onSaveDocument: () => {
            const editor = window.DocEditor?.instances[EDITOR_ID];

            console.log("onSaveDocument");
            if (editor) {
              editor.downloadAs();
            }
          },
          onDownloadAs: (event: { data: { url: string } }) => {
            Office.context.ui.messageParent(JSON.stringify({
              type: "request-save",
              data: {
                attachmentId: attachmentId.current,
                name: thisConfig.document?.title,
                url: event.data.url,
              }
            }));
          },
        }

        setDocumentServerUrl(thisDocumentServerUrl);
        setConfig(thisConfig);
        attachmentId.current = thisAttacmentId;
      } else if (message.type === "response-save") {
        attachmentId.current = message.data.attachmentId;
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
            id={EDITOR_ID}
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
