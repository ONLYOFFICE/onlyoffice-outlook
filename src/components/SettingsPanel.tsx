import * as React from "react";
import { makeStyles, tokens, Button, Field, Input } from "@fluentui/react-components";
import { ArrowLeftRegular } from "@fluentui/react-icons";
import { APP_SETTINGS_KEY, DOCUMENT_SERVER_JWT_SECRET_SETTING, DOCUMENT_SERVER_URL_SETTING } from "../constants";

/* global Office */

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
    alignItems: "center",
    gap: "8px",
  },
  title: {
    margin: "0",
    fontSize: "20px",
    fontWeight: "600",
    lineHeight: "1.3",
    color: tokens.colorNeutralForeground1,
  },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    paddingTop: "16px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  status: {
    margin: "0",
    fontSize: "12px",
    color: tokens.colorNeutralForeground3,
    minHeight: "16px",
  },
  statusError: {
    color: tokens.colorPaletteRedForeground1,
  },
});

interface SettingsPanelProps {
  onBack: () => void;
}

const SettingsPanel: React.FC<SettingsPanelProps> = ({ onBack }) => {
  const styles = useStyles();

  const loadSettings = () => {
    const appSettings = Office.context.roamingSettings.get(APP_SETTINGS_KEY) || {};
    return {
      url: (appSettings[DOCUMENT_SERVER_URL_SETTING] as string) || "",
      secret: (appSettings[DOCUMENT_SERVER_JWT_SECRET_SETTING] as string) || "",
    };
  };

  const initial = loadSettings();
  const [url, setUrl] = React.useState(initial.url);
  const [secret, setSecret] = React.useState(initial.secret);
  const [status, setStatus] = React.useState<{ text: string; error: boolean } | null>(null);

  function save() {
    const current = Office.context.roamingSettings.get(APP_SETTINGS_KEY) || {};
    current[DOCUMENT_SERVER_URL_SETTING] = url.trim();
    current[DOCUMENT_SERVER_JWT_SECRET_SETTING] = secret.trim();
    Office.context.roamingSettings.set(APP_SETTINGS_KEY, current);
    Office.context.roamingSettings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        setStatus({ text: "Settings saved.", error: false });
      } else {
        setStatus({ text: "Failed to save settings.", error: true });
      }
    });
  }

  return (
    <div className={styles.root}>
      <header className={styles.header}>
        <Button appearance="subtle" icon={<ArrowLeftRegular />} onClick={onBack} aria-label="Back" />
        <h1 className={styles.title}>Settings</h1>
      </header>

      <div className={styles.section}>
        <Field label="ONLYOFFICE Docs address">
          <Input
            type="url"
            placeholder="https://your-docs-server.com"
            value={url}
            onChange={(_e, data) => { setUrl(data.value); setStatus(null); }}
            autoComplete="off"
          />
        </Field>

        <Field label="ONLYOFFICE Docs secret key">
          <Input
            type="password"
            placeholder=""
            value={secret}
            onChange={(_e, data) => { setSecret(data.value); setStatus(null); }}
            autoComplete="off"
          />
        </Field>

        <Button appearance="primary" onClick={save} style={{ alignSelf: "flex-start" }}>
          Save
        </Button>

        {status && (
          <p className={`${styles.status}${status.error ? ` ${styles.statusError}` : ""}`}>
            {status.text}
          </p>
        )}
      </div>
    </div>
  );
};

export default SettingsPanel;
