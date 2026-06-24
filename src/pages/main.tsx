import * as React from "react";
import { makeStyles, shorthands, tokens, Button } from "@fluentui/react-components";
import {
  SettingsRegular,
  DocumentRegular,
  TableRegular,
  SlideLayoutRegular,
  DocumentPdfRegular,
  AttachRegular,
} from "@fluentui/react-icons";
import { FileUtils } from "../utils/fileUtils";
import { DocumentServerClient } from "../client/DocumentServerClient";
import { APP_SETTINGS_KEY, DOCUMENT_SERVER_JWT_SECRET_SETTING, DOCUMENT_SERVER_URL_SETTING } from "../constants";

/* global Office, window, fetch */

// --- Utility functions ---

function formatFileCount(count: number): string {
  return count === 1 ? "1 file" : `${count} files`;
}

function formatFileSize(size: number): string {
  if (!Number.isFinite(size) || size <= 0) return "Unknown size";
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

interface FileKind {
  label: string;
  color: string;
  Icon: React.ComponentType;
}

function getFileKind(fileName: string, fu: FileUtils): FileKind {
  const ext = fu.getExtension(fileName);
  if (["xls", "xlsx", "csv"].includes(ext)) {
    return { label: "Spreadsheet", color: "#107c41", Icon: TableRegular };
  }
  if (["doc", "docx", "txt", "rtf"].includes(ext)) {
    return { label: "Document", color: "#185abd", Icon: DocumentRegular };
  }
  if (ext === "pdf") {
    return { label: "PDF", color: "#c4312b", Icon: DocumentPdfRegular };
  }
  if (["ppt", "pptx"].includes(ext)) {
    return { label: "Presentation", color: "#c43e1c", Icon: SlideLayoutRegular };
  }
  return { label: "File", color: "#616161", Icon: AttachRegular };
}

function getSenderLabel(item: Office.MessageRead): string {
  if (!item.from) return "Unknown sender";
  return item.from.emailAddress || item.from.displayName || "Unknown sender";
}

function formatMessageDate(value: Date | undefined): string {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return date.toLocaleString(undefined, {
    month: "short",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

// --- Constants ---

interface DocType {
  type: string;
  label: string;
  ext: string;
  color: string;
  Icon: React.ComponentType;
}

const DOC_TYPES: DocType[] = [
  { type: "docx", label: "Document", ext: ".docx", color: "#185abd", Icon: DocumentRegular },
  { type: "xlsx", label: "Spreadsheet", ext: ".xlsx", color: "#107c41", Icon: TableRegular },
  { type: "pptx", label: "Presentation", ext: ".pptx", color: "#c43e1c", Icon: SlideLayoutRegular },
];

// --- Styles ---

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    minHeight: "100vh",
    padding: "20px",
    boxSizing: "border-box",
  },
  header: {
    display: "flex",
    alignItems: "flex-start",
    justifyContent: "space-between",
  },
  title: {
    margin: "0",
    fontSize: "20px",
    fontWeight: "600",
    lineHeight: "1.3",
    color: tokens.colorNeutralForeground1,
  },
  subtitle: {
    margin: "4px 0 0",
    color: tokens.colorNeutralForeground3,
    fontSize: "13px",
  },
  messageInfo: {
    display: "grid",
    gridTemplateColumns: "44px 1fr",
    gap: "12px",
    alignItems: "center",
    padding: "14px 0",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  avatar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: "44px",
    height: "44px",
    borderRadius: "50%",
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundInverted,
    fontSize: "20px",
    fontWeight: "600",
    flexShrink: "0",
  },
  messageContent: {
    minWidth: "0",
  },
  subject: {
    margin: "0",
    overflow: "hidden",
    fontSize: "16px",
    fontWeight: "600",
    lineHeight: "1.35",
    color: tokens.colorNeutralForeground1,
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  meta: {
    display: "flex",
    gap: "6px",
    alignItems: "center",
    marginTop: "6px",
    color: tokens.colorNeutralForeground3,
    fontSize: "13px",
    overflow: "hidden",
  },
  metaFrom: {
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
    flexShrink: "1",
    minWidth: "0",
  },
  metaDate: {
    whiteSpace: "nowrap",
    flexShrink: "0",
    marginLeft: "auto",
  },
  createTitle: {
    margin: "0 0 12px",
    fontSize: "12px",
    fontWeight: "600",
    color: tokens.colorNeutralForeground3,
    textTransform: "uppercase",
    letterSpacing: "0.05em",
  },
  createGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(3, 1fr)",
    gap: "8px",
  },
  createBtn: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "8px",
    padding: "14px 8px",
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke1),
    borderRadius: "6px",
    cursor: "pointer",
    backgroundColor: tokens.colorNeutralBackground2,
    color: tokens.colorNeutralForeground1,
    fontFamily: "inherit",
    height: "auto",
    minWidth: "0",
    width: "100%",
    ":hover": {
      ...shorthands.borderColor(tokens.colorBrandStroke1),
      backgroundColor: tokens.colorBrandBackground2,
    },
    ":focus-visible": {
      ...shorthands.outline("2px", "solid", tokens.colorBrandStroke1),
      outlineOffset: "1px",
    },
  },
  createIconBox: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: "36px",
    height: "36px",
    borderRadius: "6px",
    color: "#ffffff",
    fontSize: "18px",
  },
  createLabel: {
    fontSize: "12px",
    fontWeight: "600",
    lineHeight: "1.2",
  },
  createExt: {
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
  },
  fileList: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  fileEmpty: {
    margin: "0",
    padding: "16px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: "6px",
    backgroundColor: tokens.colorNeutralBackground2,
    color: tokens.colorNeutralForeground3,
    fontSize: "13px",
  },
  fileItem: {
    display: "grid",
    gridTemplateColumns: "36px 1fr auto",
    alignItems: "center",
    gap: "12px",
    padding: "12px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: "6px",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  fileIconBox: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: "32px",
    height: "32px",
    borderRadius: "4px",
    color: "#ffffff",
    fontSize: "16px",
    flexShrink: "0",
  },
  fileContent: {
    minWidth: "0",
  },
  fileName: {
    margin: "0",
    overflow: "hidden",
    fontSize: "14px",
    fontWeight: "600",
    lineHeight: "1.3",
    color: tokens.colorNeutralForeground1,
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  fileMeta: {
    margin: "2px 0 0",
    color: tokens.colorNeutralForeground3,
    fontSize: "12px",
  },
});

// --- Main component ---

const MainPage: React.FC = () => {
  const styles = useStyles();
  const fileUtils = React.useRef(new FileUtils([]));
  const [attachments, setAttachments] = React.useState<Office.AttachmentDetails[]>([]);
  const [subject, setSubject] = React.useState("Loading message...");
  const [from, setFrom] = React.useState("Unknown sender");
  const [date, setDate] = React.useState("");
  const [fileCount, setFileCount] = React.useState("Loading files...");
  const [showCreate, setShowCreate] = React.useState(false);

  const appSettings = Office.context.roamingSettings.get(APP_SETTINGS_KEY);

  React.useEffect(() => {
    new DocumentServerClient().getFormats().then((formats) => {
      fileUtils.current = new FileUtils(formats);
    });

    const item = Office.context?.mailbox?.item;
    if (!item) {
      setFileCount("No mailbox context");
      return;
    }

    const isCompose =
      !!(item as Office.MessageCompose).subject &&
      typeof (item as Office.MessageCompose).subject.getAsync === "function";

    const canCreate =
      typeof (item as Office.MessageCompose).addFileAttachmentFromBase64Async === "function";

    if (isCompose) {
      (item as Office.MessageCompose).subject.getAsync((result) => {
        setSubject(
          result.status === Office.AsyncResultStatus.Succeeded
            ? result.value || "(No subject)"
            : "(No subject)"
        );
      });
      setFrom("Draft");
    } else {
      const readItem = item as Office.MessageRead;
      setSubject(readItem.subject || "(No subject)");
      setFrom(getSenderLabel(readItem));
      setDate(formatMessageDate(readItem.dateTimeCreated));
    }

    const atts: Office.AttachmentDetails[] = (item as Office.MessageRead).attachments || [];
    setAttachments(atts);
    setFileCount(formatFileCount(atts.length));
    setShowCreate(canCreate);
  }, []);

  function openSettings() {
    window.location.href = "settings.html";
  }

  function openEditor(attachment: Office.AttachmentDetails) {
    const editorUrl = `${window.location.origin}/index.html#editor`;

    Office.context.ui.displayDialogAsync(editorUrl, { height: 80, width: 80 }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(`Error opening editor: ${result.error.message}`);
        return;
      }

      const dialog = result.value;

      dialog.addEventHandler(
        Office.EventType.DialogMessageReceived,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (arg: any) => {
          const message = JSON.parse(arg.message);

          if (message.type === "request-config") {
            (Office.context.mailbox.item as Office.MessageRead).getAttachmentContentAsync(
              attachment.id,
              (contentResult) => {
                if (contentResult.status !== Office.AsyncResultStatus.Succeeded) {
                  console.error(`Error loading attachment: ${contentResult.error.message}`);
                  dialog.close();
                  return;
                }

                const documentServerUrl = appSettings ? appSettings[DOCUMENT_SERVER_URL_SETTING] || "https://3998-3-125-222-163.ngrok-free.app" : "https://3998-3-125-222-163.ngrok-free.app";
                const documentServerJwtSecret = appSettings ? appSettings[DOCUMENT_SERVER_JWT_SECRET_SETTING] || "EPAvpORzhQ1lNB6PeTPWGD4MgX7w6MyJ" : "EPAvpORzhQ1lNB6PeTPWGD4MgX7w6MyJ";

                dialog.messageChild(
                  JSON.stringify({
                    type: "response-config",
                    data: {
                      documentServerUrl,
                      config: fileUtils.current.createEditorConfig(
                        Date.now().toString(),
                        attachment.name,
                        "_data_",
                        "view",
                        documentServerJwtSecret
                      ),
                      content: contentResult.value.content
                    },
                  })
                );
              }
            );
          }
        }
      );

      dialog.addEventHandler(
        Office.EventType.DialogEventReceived,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (_arg: any) => {}
      );
    });
  }

  function openNewDocumentInEditor(docType: string) {
    const editorUrl = `${window.location.origin}/editor.html?mode=create&fileType=${encodeURIComponent(docType)}&fileName=Document`;

    Office.context.ui.displayDialogAsync(editorUrl, { height: 80, width: 80 }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(`Error opening editor: ${result.error.message}`);
        return;
      }

      const dialog = result.value;

      dialog.addEventHandler(
        Office.EventType.DialogMessageReceived,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (arg: any) => {
          const message = JSON.parse(arg.message);

          if (message.type === "create-attachment") {
            const proxyUrl = `/download-proxy?url=${encodeURIComponent(message.data.url)}`;
            fetch(proxyUrl)
              .then((r) => r.json())
              .then(({ base64 }: { base64: string }) => {
                (Office.context.mailbox.item as Office.MessageCompose).addFileAttachmentFromBase64Async(
                  base64,
                  `Document.${message.data.fileType}`,
                  (asyncResult) => {
                    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                      console.error("Error adding attachment:", asyncResult.error.message);
                    }
                  }
                );
              })
              .catch((err: Error) => console.error("Error loading file:", err));
          }
        }
      );

      dialog.addEventHandler(
        Office.EventType.DialogEventReceived,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (_arg: any) => {}
      );
    });
  }

  const senderInitial = from ? from.charAt(0).toUpperCase() : "M";

  return (
    <div className={styles.root}>
      <header className={styles.header}>
        <div>
          <h1 className={styles.title}>ONLYOFFICE Attachments</h1>
          <p className={styles.subtitle}>{fileCount}</p>
        </div>
        <Button
          appearance="subtle"
          icon={<SettingsRegular />}
          onClick={openSettings}
          aria-label="Settings"
        />
      </header>

      <section className={styles.messageInfo} aria-label="Message details">
        <div className={styles.avatar} aria-hidden="true">
          {senderInitial}
        </div>
        <div className={styles.messageContent}>
          <h2 className={styles.subject}>{subject}</h2>
          <div className={styles.meta}>
            <span>From:</span>
            <span className={styles.metaFrom}>{from}</span>
            {date && <time className={styles.metaDate}>{date}</time>}
          </div>
        </div>
      </section>

      {showCreate && (
        <section aria-label="Create document">
          <h2 className={styles.createTitle}>Create document</h2>
          <div className={styles.createGrid}>
            {DOC_TYPES.map((doc) => (
              <button
                key={doc.type}
                className={styles.createBtn}
                type="button"
                onClick={() => openNewDocumentInEditor(doc.type)}
              >
                <div
                  className={styles.createIconBox}
                  style={{ backgroundColor: doc.color }}
                  aria-hidden="true"
                >
                  <doc.Icon />
                </div>
                <span className={styles.createLabel}>{doc.label}</span>
                <span className={styles.createExt}>{doc.ext}</span>
              </button>
            ))}
          </div>
        </section>
      )}

      <section className={styles.fileList} aria-label="Files">
        {attachments.length === 0 ? (
          <p className={styles.fileEmpty}>No attachments in this message.</p>
        ) : (
          attachments.map((att) => {
            const kind = getFileKind(att.name || "", fileUtils.current);
            const fileName = att.name || "Unnamed attachment";
            return (
              <article key={att.id} className={styles.fileItem}>
                <div
                  className={styles.fileIconBox}
                  style={{ backgroundColor: kind.color }}
                  aria-hidden="true"
                >
                  <kind.Icon />
                </div>
                <div className={styles.fileContent}>
                  <h2 className={styles.fileName}>{fileName}</h2>
                  <p className={styles.fileMeta}>
                    {kind.label} · {formatFileSize(att.size)}
                  </p>
                </div>
                <Button
                  appearance="primary"
                  size="small"
                  onClick={() => openEditor(att)}
                >
                  View
                </Button>
              </article>
            );
          })
        )}
      </section>

    </div>
  );
};

export default MainPage;
