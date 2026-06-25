import * as React from "react";
import {
  makeStyles,
  shorthands,
  tokens,
  Button,
  Spinner,
  Field,
  Input,
  Avatar,
} from "@fluentui/react-components";
import {
  SettingsRegular,
  EditRegular,
  EyeRegular,
  ArrowDownloadRegular,
  DeleteRegular,
} from "@fluentui/react-icons";
import SettingsPanel from "../components/SettingsPanel";
import { FileUtils } from "../utils/fileUtils";
import { DocumentServerClient } from "../client/DocumentServerClient";
import {
  APP_SETTINGS_KEY,
  DOCUMENT_SERVER_JWT_SECRET_SETTING,
  DOCUMENT_SERVER_URL_SETTING,
} from "../constants";

/* global Office, window, crypto, console, document */

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

function formatEmailAddress(details: Office.EmailAddressDetails | null | undefined): string {
  if (!details) return "";
  const { displayName, emailAddress } = details;
  if (displayName && emailAddress && displayName !== emailAddress) {
    return `${displayName}<${emailAddress}>`;
  }
  return emailAddress || displayName || "";
}

function getSenderLabel(item: Office.MessageRead): string {
  return formatEmailAddress(item.from) || "Unknown sender";
}

function getRecipientsLabel(recipients: Office.EmailAddressDetails[]): string {
  if (!recipients || recipients.length === 0) return "";
  return recipients.map(formatEmailAddress).filter(Boolean).join(", ");
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
  icon: string;
}

const ONLYOFFICE_BLANK_FILE_URL = "https://static.onlyoffice.com/assets/docs/samples/blank";

const DOC_TYPES: DocType[] = [
  { type: "docx", label: "Document", icon: "assets/word.svg" },
  { type: "xlsx", label: "Spreadsheet", icon: "assets/cell.svg" },
  { type: "pptx", label: "Presentation", icon: "assets/slide.svg" },
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
    flexDirection: "column",
    gap: "4px",
    marginTop: "6px",
    color: tokens.colorNeutralForeground3,
    fontSize: "13px",
    overflow: "hidden",
  },
  metaRow: {
    display: "flex",
    gap: "6px",
    alignItems: "center",
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
  createForm: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    marginTop: "12px",
    padding: "12px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: "6px",
    backgroundColor: tokens.colorNeutralBackground2,
  },
  createFormActions: {
    display: "flex",
    gap: "8px",
    justifyContent: "flex-end",
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
  actionSpinner: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: "28px",
    height: "28px",
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
};

const canEdit = (item: Office.Item) => {
  return canCreate(item) && canDelete(item);
};

const MainPage: React.FC = () => {
  const styles = useStyles();
  const [page, setPage] = React.useState<"main" | "settings">("main");
  const [fileUtils, setFileUtils] = React.useState<FileUtils>();
  const [attachments, setAttachments] = React.useState<
    Office.AttachmentDetails[] | Office.AttachmentDetailsCompose[]
  >([]);
  const [subject, setSubject] = React.useState("Loading message...");
  const [from, setFrom] = React.useState("Unknown sender");
  const [senderDisplayName, setSenderDisplayName] = React.useState("Unknown sender");
  const [to, setTo] = React.useState("");
  const [date, setDate] = React.useState("");
  const [fileCount, setFileCount] = React.useState("Loading files...");
  const [showCreate, setShowCreate] = React.useState(false);
  const [downloading, setDownloading] = React.useState<Set<string>>(new Set());
  const [deleting, setDeleting] = React.useState<Set<string>>(new Set());
  const [pendingCreateType, setPendingCreateType] = React.useState<string | null>(null);
  const [pendingCreateName, setPendingCreateName] = React.useState("");

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
      (item as Office.MessageCompose).from.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          setFrom(formatEmailAddress(result.value) || "Draft");
          setSenderDisplayName(result.value?.displayName || result.value?.emailAddress || "Draft");
        }
      });
      (item as Office.MessageCompose).to.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          setTo(getRecipientsLabel(result.value));
        }
      });
    } else {
      const readItem = item as Office.MessageRead;
      setSubject(readItem.subject || "(No subject)");
      setFrom(getSenderLabel(readItem));
      setSenderDisplayName(
        readItem.from?.displayName || readItem.from?.emailAddress || "Unknown sender"
      );
      setTo(getRecipientsLabel(readItem.to));
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

  async function createEditorConfig(
    attachmentId: string,
    attachmentName: string,
    fileUrl: string,
    content: string | undefined = undefined
  ) {
    if (!fileUtils) throw new Error("FileUtils not initialized");

    const appSettings = Office.context.roamingSettings.get(APP_SETTINGS_KEY) || {};
    const documentServerUrl = (appSettings[DOCUMENT_SERVER_URL_SETTING] as string) || "";
    const documentServerJwtSecret =
      (appSettings[DOCUMENT_SERVER_JWT_SECRET_SETTING] as string) || "";

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
  }

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
            (Office.context.mailbox.item as Office.MessageCompose).addFileAttachmentFromBase64Async(
              message.data.data,
              message.data.name,
              function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  dialog.messageChild(
                    JSON.stringify({
                      type: "response-save",
                      data: {
                        attachmentId: asyncResult.value,
                      },
                    })
                  );
                  if (message.data.attachmentId) {
                    (Office.context.mailbox.item as Office.MessageCompose).removeAttachmentAsync(
                      message.data.attachmentId
                    );
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
    });
  }

  function downloadAttachment(
    attachment: Office.AttachmentDetails | Office.AttachmentDetailsCompose
  ) {
    if (!attachment.id) return;

    setDownloading((prev) => new Set(prev).add(attachment.id));

    (Office.context.mailbox.item as Office.MessageRead).getAttachmentContentAsync(
      attachment.id,
      (result: Office.AsyncResult<Office.AttachmentContent>) => {
        setDownloading((prev) => {
          const next = new Set(prev);
          next.delete(attachment.id);
          return next;
        });
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

  function deleteAttachment(
    attachment: Office.AttachmentDetails | Office.AttachmentDetailsCompose
  ) {
    if (!attachment.id) return;

    setDeleting((prev) => new Set(prev).add(attachment.id));

    (Office.context.mailbox.item as Office.MessageCompose).removeAttachmentAsync(
      attachment.id,
      (result) => {
        setDeleting((prev) => {
          const next = new Set(prev);
          next.delete(attachment.id);
          return next;
        });
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.error(`Error deleting attachment: ${result.error.message}`);
        }
      }
    );
  }

  function onCreateNew(type: string, name: string) {
    const attachment = {
      id: null,
      name: name + "." + type,
    } as unknown as Office.AttachmentDetails;

    openEditor(attachment);
  }

  function submitCreateForm() {
    if (!pendingCreateType) return;
    const defaultLabel = DOC_TYPES.find((d) => d.type === pendingCreateType)?.label ?? "Document";
    const name = pendingCreateName.trim() || defaultLabel;
    setPendingCreateType(null);
    setPendingCreateName("");
    onCreateNew(pendingCreateType, name);
  }

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
          onClick={() => {
            setPage("settings");
          }}
          aria-label="Settings"
        />
      </header>

      <section className={styles.messageInfo} aria-label="Message details">
        <Avatar color="colorful" name={senderDisplayName} size={40} aria-hidden="true" />
        <div className={styles.messageContent}>
          <h2 className={styles.subject}>{subject}</h2>
          <div className={styles.meta}>
            <div className={styles.metaRow}>
              <span>From:</span>
              <span className={styles.metaFrom}>{from}</span>
              {date && <time className={styles.metaDate}>{date}</time>}
            </div>
            {to && (
              <div className={styles.metaRow}>
                <span>To:</span>
                <span className={styles.metaFrom}>{to}</span>
              </div>
            )}
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
                onClick={() => {
                  setPendingCreateType(doc.type);
                  setPendingCreateName("");
                }}
              >
                <div className={styles.createIconBox} aria-hidden="true">
                  <img src={doc.icon} width="32" height="32" alt="" />
                </div>
                <span className={styles.createLabel}>{doc.label}</span>
              </button>
            ))}
          </div>
          {pendingCreateType && (
            <div className={styles.createForm}>
              <Field label="File name">
                <Input
                  autoFocus
                  placeholder={
                    DOC_TYPES.find((d) => d.type === pendingCreateType)?.label ?? "Document"
                  }
                  value={pendingCreateName}
                  onChange={(_, data) => setPendingCreateName(data.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") submitCreateForm();
                    if (e.key === "Escape") {
                      setPendingCreateType(null);
                      setPendingCreateName("");
                    }
                  }}
                  contentAfter={
                    <span style={{ color: "var(--colorNeutralForeground3)", fontSize: "12px" }}>
                      .{pendingCreateType}
                    </span>
                  }
                />
              </Field>
              <div className={styles.createFormActions}>
                <Button
                  appearance="secondary"
                  size="small"
                  onClick={() => {
                    setPendingCreateType(null);
                    setPendingCreateName("");
                  }}
                >
                  Cancel
                </Button>
                <Button appearance="primary" size="small" onClick={submitCreateForm}>
                  Create
                </Button>
              </div>
            </div>
          )}
        </section>
      )}

      <section className={styles.fileList} aria-label="Files">
        {!fileUtils
          ? null
          : (() => {
              const viewable = attachments.filter((att) =>
                fileUtils.isViewable(fileUtils.getExtension(att.name || ""))
              );
              if (viewable.length === 0)
                return <p className={styles.fileEmpty}>No attachments in this message.</p>;
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
                      {canEdit(Office.context.mailbox.item || {}) &&
                        fileUtils.isEditable(extension) && (
                          <Button
                            appearance="subtle"
                            size="small"
                            icon={<EditRegular />}
                            onClick={() => openEditor(att)}
                            aria-label="Edit"
                          />
                        )}
                      {!(
                        canEdit(Office.context.mailbox.item || {}) &&
                        fileUtils.isEditable(extension)
                      ) && (
                        <Button
                          appearance="subtle"
                          size="small"
                          icon={<EyeRegular />}
                          onClick={() => openEditor(att)}
                          aria-label="View"
                        />
                      )}
                      {downloading.has(att.id) ? (
                        <div className={styles.actionSpinner}>
                          <Spinner size="extra-tiny" />
                        </div>
                      ) : (
                        <Button
                          appearance="subtle"
                          size="small"
                          icon={<ArrowDownloadRegular />}
                          onClick={() => downloadAttachment(att)}
                          aria-label="Download"
                        />
                      )}
                      {canDelete(Office.context.mailbox.item || {}) &&
                        (deleting.has(att.id) ? (
                          <div className={styles.actionSpinner}>
                            <Spinner size="extra-tiny" />
                          </div>
                        ) : (
                          <Button
                            appearance="subtle"
                            size="small"
                            icon={<DeleteRegular />}
                            onClick={() => deleteAttachment(att)}
                            aria-label="Delete"
                          />
                        ))}
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
