import { makeStyles, Spinner, Text, Button } from "@fluentui/react-components";
import { DocumentEditor } from "@onlyoffice/document-editor-react";
import * as React from "react";

import { Config } from '@onlyoffice/doceditor-types';

const base64ToUint8Array = (base64: string) => {
  const binary = atob(base64);

  return new Uint8Array(Array.from(binary, (char) => char.charCodeAt(0)));
}

const uint8ArrayToBase64 = (buffer: ArrayBuffer | Uint8Array): string => {
  const uint8Array = buffer instanceof Uint8Array 
    ? buffer 
    : new Uint8Array(buffer);
    
  return btoa(Array.from(uint8Array, (byte) => String.fromCharCode(byte)).join(''));
}

const useStyles = makeStyles({
  root: {
    height: "100%",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    textAlign: "center",
  },
  errorContainer: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    gap: "12px",
    padding: "32px",
    maxWidth: "420px",
  },
  errorIcon: {
    fontSize: "48px",
    lineHeight: "1",
    marginBottom: "8px",
  },
  errorTitle: {
    fontWeight: "600",
  },
  errorDescription: {
    color: "#605e5c",
    lineHeight: "1.5",
  },
  savingOverlay: {
    position: "fixed",
    inset: "0",
    backgroundColor: "rgba(0, 0, 0, 0.4)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    zIndex: "9999",
    pointerEvents: "all",
  },
  savingOverlayHidden: {
    display: "none",
  },
});

const EDITOR_ID = "onlyofficeEditor";

const EditorPage: React.FC = () => {
  const styles = useStyles();
  const [documentServerUrl, setDocumentServerUrl] = React.useState();
  const [config, setConfig] = React.useState<Config | undefined>();
  const [isLoading, setIsLoading] = React.useState(true);
  const [isSaving, setIsSaving] = React.useState(false);
  const [hasError, setHasError] = React.useState(false);
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

            if (editor && content && thisConfig.document.url === "_data_") {
              const body = base64ToUint8Array(content);
              //@ts-expect-error
              editor.openDocument(body);
            }
          },
          onSaveDocument: (event: any) => {
            setIsSaving(true);

            Office.context.ui.messageParent(JSON.stringify({
              type: "request-save",
              data: {
                attachmentId: attachmentId.current,
                name: thisConfig.document?.title,
                data: uint8ArrayToBase64(event.data),
              }
            }));
          },
        }

        setDocumentServerUrl(thisDocumentServerUrl);
        setConfig(thisConfig);
        attachmentId.current = thisAttacmentId;
        setIsLoading(false);
      } else if (message.type === "response-save") {
        const editor = window.DocEditor?.instances[EDITOR_ID];

        attachmentId.current = message.data.attachmentId;
        setIsSaving(false);
        if (editor) {
          editor.showMessage("Document saved successfully.");
        }
      }
    }, () => {});

    Office.context.ui.messageParent(JSON.stringify({
      type: "request-config"
    }));
  }, []);

  const onLoadComponentError = () => {
    setHasError(true);
    setIsLoading(false);
  };

  return (
    <div className={styles.root}>
      <div className={isSaving ? styles.savingOverlay : styles.savingOverlayHidden}>
        <Spinner size="large" label="Saving..." labelPosition="below" />
      </div>
      {isLoading
        ? <Spinner size="large" label="Opening document..." labelPosition="below" />
        : hasError
          ? <div className={styles.errorContainer}>
              <div className={styles.errorIcon}>⚠️</div>
              <Text size={500} className={styles.errorTitle}>Document Server Unavailable</Text>
              <Text size={300} className={styles.errorDescription}>
                Unable to connect to ONLYOFFICE Document Server.
                Please check your add-in settings and ensure the server address is correct and the server is running.
              </Text>
              <Button appearance="primary" onClick={() => Office.context.ui.messageParent(JSON.stringify({ type: "request-open-settings" }))}>
                Open Settings
              </Button>
            </div>
          : documentServerUrl && config &&
            <DocumentEditor
                id={EDITOR_ID}
                documentServerUrl={documentServerUrl}
                config={config}
                height="100%"
                width="100%"
                onLoadComponentError={onLoadComponentError}
              />
      }
    </div>
  );
};

export default EditorPage;
