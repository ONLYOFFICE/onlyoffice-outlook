import * as React from "react";
import {
  makeStyles,
  shorthands,
  tokens,
  Button,
} from "@fluentui/react-components";
import {
  SettingsRegular,
  DocumentRegular,
  TableRegular,
  SlideLayoutRegular,
  EditRegular,
  EyeRegular,
  ArrowDownloadRegular,
  DeleteRegular,
} from "@fluentui/react-icons";
import SettingsPanel from "../components/SettingsPanel";
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

const DOC_TYPE_LABELS: Record<string, string> = {
  word: "Document",
  cell: "Spreadsheet",
  slide: "Presentation",
  pdf: "PDF",
  diagram: "Diagram",
};

const DOC_TYPE_ICONS: Record<string, string> = {
  word: "assets/word.svg",
  cell: "assets/cell.svg",
  slide: "assets/slide.svg",
  pdf: "assets/pdf.svg",
  diagram: "assets/diagram.svg",
};

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
  color: string;
  Icon: React.ComponentType;
}

const ONLYOFFICE_BLANK_FILE_URL = "https://static.onlyoffice.com/assets/docs/samples/blank";

const DOC_TYPES: DocType[] = [
  { type: "docx", label: "Document", color: "#185abd", Icon: DocumentRegular },
  { type: "xlsx", label: "Spreadsheet", color: "#107c41", Icon: TableRegular },
  { type: "pptx", label: "Presentation", color: "#c43e1c", Icon: SlideLayoutRegular },
];

// --- Styles ---

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    height: "100vh",
    padding: "20px",
    boxSizing: "border-box",
    overflow: "hidden",
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
  fileActions: {
    display: "flex",
    gap: "2px",
    alignItems: "center",
    flexShrink: "0",
  },
  fileList: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    flex: "1",
    overflowY: "auto",
    minHeight: "0",
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

const canCreate = (item: Office.Item) => {
  return typeof (item as Office.MessageCompose).addFileAttachmentFromBase64Async === "function";
};

const canDelete = (item: Office.Item) => {
  return typeof (item as Office.MessageCompose).removeAttachmentAsync === "function";
}

const canEdit = (item: Office.Item) => {
  return canCreate(item) && canDelete(item);
};

const MainPage: React.FC = () => {
  const styles = useStyles();
  const [page, setPage] = React.useState<"main" | "settings">("main");
  const [fileUtils, setFileUtils] = React.useState<FileUtils>();
  const [attachments, setAttachments] = React.useState<Office.AttachmentDetails[] | Office.AttachmentDetailsCompose[]>([]);
  const [subject, setSubject] = React.useState("Loading message...");
  const [from, setFrom] = React.useState("Unknown sender");
  const [date, setDate] = React.useState("");
  const [fileCount, setFileCount] = React.useState("Loading files...");
  const [showCreate, setShowCreate] = React.useState(false);

  React.useEffect(() => {
    new DocumentServerClient().getFormats().then((formats) => {
      setFileUtils(new FileUtils(formats));
    });
  }, []);

  React.useEffect(() => {
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

    setShowCreate(canCreate);

    if (isCompose) {
      (item as Office.MessageCompose).addHandlerAsync(Office.EventType.AttachmentsChanged, () => {
        loadAttachments();
      });
    }

    loadAttachments();
  }, []);

  const loadAttachments = () => {
    const item = Office.context.mailbox.item;

    if (typeof (item as Office.MessageCompose).getAttachmentsAsync === "function") {
      (item as Office.MessageCompose).getAttachmentsAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          setAttachments(result.value);
          setFileCount(formatFileCount(result.value.length));
        }
      });
    } else {
      setAttachments((item as Office.MessageRead).attachments);
      setFileCount(formatFileCount((item as Office.MessageRead).attachments.length));
    }
  };

  if (page === "settings") {
    return <SettingsPanel onBack={() => setPage("main")} />;
  }

  async function createEditorConfig(attachmentId: string, attachmentName: string, fileUrl: string, content: string | undefined = undefined) {
    if (!fileUtils) throw new Error("FileUtils not initialized");

    const appSettings = Office.context.roamingSettings.get(APP_SETTINGS_KEY) || {};
    const documentServerUrl = (appSettings[DOCUMENT_SERVER_URL_SETTING] as string) || "";
    const documentServerJwtSecret = (appSettings[DOCUMENT_SERVER_JWT_SECRET_SETTING] as string) || "";

    const mode = canEdit(Office.context.mailbox.item || {}) ? "edit" : "view";
    const key = attachmentId ? await fileUtils.createKey(attachmentId) : crypto.randomUUID();
    const user = {
      id: Office.context.mailbox.userProfile.emailAddress,
      name: Office.context.mailbox.userProfile.displayName,
    };
    const locale = Office.context.displayLanguage;

    return {
      documentServerUrl,
      config: fileUtils.createEditorConfig(
        key,
        attachmentName,
        fileUrl,
        mode,
        user,
        documentServerJwtSecret,
        locale
      ),
      content,
      attachmentId,
    };
  };

  function openEditor(attachment: Office.AttachmentDetails | Office.AttachmentDetailsCompose) {
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
        async (arg: any) => {
          const message = JSON.parse(arg.message);

          if (message.type === "request-config") {
            if (attachment.id) {
              (Office.context.mailbox.item as Office.MessageCompose).getAttachmentContentAsync(
                attachment.id,
                async (contentResult) => {
                  if (contentResult.status !== Office.AsyncResultStatus.Succeeded) {
                    console.error(`Error loading attachment: ${contentResult.error.message}`);
                    dialog.close();
                    return;
                  }

                  dialog.messageChild(
                    JSON.stringify({
                      type: "response-config",
                      data: await createEditorConfig(
                        attachment.id,
                        attachment.name,
                        "_data_",
                        contentResult.value.content
                      ),
                    })
                  );
                }
              );
            } else {
              dialog.messageChild(
                JSON.stringify({
                  type: "response-config",
                  data: await createEditorConfig(
                    attachment.id,
                    attachment.name,
                    ONLYOFFICE_BLANK_FILE_URL
                  ),
                })
              );
            }
          } else if (message.type === "request-save") {
            const proxyUrl = `${window.location.origin}/api/proxy?url=${encodeURIComponent(message.data.url)}`;
            (Office.context.mailbox.item as Office.MessageCompose).addFileAttachmentAsync(
              proxyUrl,
              message.data.name,
              function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  dialog.messageChild(JSON.stringify({
                    type: "response-save",
                    data: {
                      attachmentId: asyncResult.value,
                    }
                  }));
                  if (message.data.attachmentId) {
                    (Office.context.mailbox.item as Office.MessageCompose).removeAttachmentAsync(message.data.attachmentId);
                  }
                } else {
                  console.error("Error:", asyncResult.error.message);
                }
              }
            );
          } else if (message.type === "request-open-settings") {
            dialog.close();
            setPage("settings");
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

  function downloadAttachment(attachment: Office.AttachmentDetails | Office.AttachmentDetailsCompose) {
    if (!attachment.id) return;

    (Office.context.mailbox.item as Office.MessageRead).getAttachmentContentAsync(
      attachment.id,
      (result: Office.AsyncResult<Office.AttachmentContent>) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) return;
        const link = document.createElement("a");
        link.href = `data:application/octet-stream;base64,${result.value.content}`;
        link.download = attachment.name;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
      }
    );
  }

  function deleteAttachment(attachment: Office.AttachmentDetails | Office.AttachmentDetailsCompose) {
    if (!attachment.id) return;
    (Office.context.mailbox.item as Office.MessageCompose).removeAttachmentAsync(
      attachment.id,
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.error(`Error deleting attachment: ${result.error.message}`);
        }
      }
    );
  }

  function onCreateNew(type: string) {
    const attachment = {
      id: null,
      name: "Document." + type
    } as unknown as Office.AttachmentDetails;

    openEditor(attachment);
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
          onClick={() => {setPage("settings");}}
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
                onClick={() => onCreateNew(doc.type)}
              >
                <div
                  className={styles.createIconBox}
                  style={{ backgroundColor: doc.color }}
                  aria-hidden="true"
                >
                  <doc.Icon />
                </div>
                <span className={styles.createLabel}>{doc.label}</span>
              </button>
            ))}
          </div>
        </section>
      )}

      <section className={styles.fileList} aria-label="Files">
        {!fileUtils ? null :(() => {
          const viewable = attachments.filter((att) => fileUtils.isViewable(fileUtils.getExtension(att.name || "")));
          if (viewable.length === 0) return <p className={styles.fileEmpty}>No attachments in this message.</p>;
          return viewable.map((att) => {
            const fileName = att.name || "Unnamed attachment";
            const extension = fileUtils.getExtension(fileName);
            const docType = fileUtils.getDocumentType(extension) || "";
            const iconSrc = DOC_TYPE_ICONS[docType];
            const label = DOC_TYPE_LABELS[docType] || "File";
            return (
              <article key={att.id} className={styles.fileItem}>
                <div className={styles.fileIconBox} aria-hidden="true">
                  {iconSrc && <img src={iconSrc} width="32" height="32" alt="" />}
                </div>
                <div className={styles.fileContent}>
                  <h2 className={styles.fileName}>{fileName}</h2>
                  <p className={styles.fileMeta}>
                    {label} · {formatFileSize(att.size)}
                  </p>
                </div>
                <div className={styles.fileActions}>
                  {canEdit(Office.context.mailbox.item || {}) && fileUtils.isEditable(extension) && (
                    <Button
                      appearance="subtle"
                      size="small"
                      icon={<EditRegular />}
                      onClick={() => openEditor(att)}
                      aria-label="Edit"
                    />
                  )}
                  {!(canEdit(Office.context.mailbox.item || {}) && fileUtils.isEditable(extension)) && (
                  <Button
                    appearance="subtle"
                    size="small"
                    icon={<EyeRegular />}
                    onClick={() => openEditor(att)}
                    aria-label="View"
                  />
                  )}
                  <Button
                    appearance="subtle"
                    size="small"
                    icon={<ArrowDownloadRegular />}
                    onClick={() => downloadAttachment(att)}
                    aria-label="Download"
                  />
                  {canDelete(Office.context.mailbox.item || {}) && (
                    <Button
                      appearance="subtle"
                      size="small"
                      icon={<DeleteRegular />}
                      onClick={() => deleteAttachment(att)}
                      aria-label="Delete"
                    />
                  )}
                </div>
              </article>
            );
          });
        })()}
      </section>

    </div>
  );
};

export default MainPage;
