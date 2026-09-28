import React, { useEffect, useState } from "react";
import {
  Switch,
  Dropdown,
  Option,
  Field,
  Button,
  makeStyles,
  tokens,
} from "@fluentui/react-components";

/* global Office */

const SETTINGS_KEY = "docToolkitConfig";
const DEFAULT_CONFIG = {
  defaultFont: "Calibri",
  autoApplyOnSelect: false,
  theme: "light",
};

const useStyles = makeStyles({
  section: { display: "flex", flexDirection: "column", gap: "12px" },
  ContentWrapper: {
    paddingLeft: "10px",
  },
  contentCover: {
    paddingLeft: "25px",
  },
  heading: {
    color: "white",
    fontSize: "20px",
  },
  label: {
    marginBottom: "10px",
  },
  wrapButtonSize: {
    display: "flex",
    justifyContent: "flex-end",
    marginRight: "40px",
    maxWidth: "250px",
    minWidth: "250px",
  },
  messageBar: {
    marginTop: "10px",
  },
  buttonRow: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    flexWrap: "wrap",
  },
  buttonSave: {
    marginTop: "10px",
  },
});

export default function ConfigSettings() {
  const styles = useStyles();
  const [config, setConfig] = useState(DEFAULT_CONFIG);
  const [saved, setSaved] = useState(false);

  useEffect(() => {
    const stored = Office.context.document.settings.get(SETTINGS_KEY);
    if (stored) setConfig({ ...DEFAULT_CONFIG, ...stored });
  }, []);

  function update(patch) {
    setConfig((prev) => ({ ...prev, ...patch }));
    setSaved(false);
  }

  function saveSettings() {
    Office.context.document.settings.set(SETTINGS_KEY, config);
    Office.context.document.settings.saveAsync((result) => {
      setSaved(result.status === Office.AsyncResultStatus.Succeeded);
    });
  }

  return (
    <div className={styles.ContentWrapper}>
      <h2 className={styles.heading}>16) Settings</h2>

      <div className={styles.contentCover}>
        <Field label="Default font for new formatting actions">
          <Dropdown
            value={config.defaultFont}
            onOptionSelect={(_, data) => update({ defaultFont: data.optionValue })}
          >
            {["Calibri", "Arial", "Times New Roman", "Georgia"].map((f) => (
              <Option key={f} value={f}>
                {f}
              </Option>
            ))}
          </Dropdown>
        </Field>

        <Switch
          label="Auto-apply default font when text is selected"
          checked={config.autoApplyOnSelect}
          onChange={(_, data) => update({ autoApplyOnSelect: data.checked })}
        />

        <Field label="Task pane theme">
          <Dropdown
            value={config.theme}
            onOptionSelect={(_, data) => update({ theme: data.optionValue })}
          >
            <Option value="light">Light</Option>
            <Option value="dark">Dark</Option>
          </Dropdown>
        </Field>

        <div className={styles.label}>
          <Button appearance="primary" onClick={saveSettings} className={styles.buttonSave}>
            Save settings
          </Button>
        </div>
        {saved && <span>Settings saved to this document.</span>}
      </div>
    </div>
  );
}
