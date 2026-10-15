import React, { useState } from "react";
import { Button, Input, Field } from "@fluentui/react-components";
import { setReportHeader, setReportFooter } from "./utils/header-footer-utils";
import { StepBody, Hint, SubHeading, Group, Actions, Status } from "../ui/StepUI";

/* global Word */

export default function HeaderFooter() {
  const [authorName, setAuthorName] = useState("");
  const [shortTitle, setShortTitle] = useState("");
  const [subjectCode, setSubjectCode] = useState("FYP4132");
  const [headerMsg, setHeaderMsg] = useState({ text: "", isError: false });
  const [footerMsg, setFooterMsg] = useState({ text: "", isError: false });

  async function handleSetHeader() {
    if (!authorName.trim() || !shortTitle.trim()) {
      setHeaderMsg({ text: "Type your name and the FYP title first.", isError: true });
      return;
    }
    try {
      await Word.run(async (context) => {
        await setReportHeader(context, authorName.trim(), shortTitle.trim());
      });
      setHeaderMsg({
        text: "Header set from section 3 onward (not on the cover pages).",
        isError: false,
      });
    } catch (err) {
      console.error(err);
      setHeaderMsg({ text: err.message || "Could not set the header.", isError: true });
    }
  }

  async function handleSetFooter() {
    if (!subjectCode.trim()) {
      setFooterMsg({ text: "Type the subject code first.", isError: true });
      return;
    }
    try {
      await Word.run(async (context) => {
        await setReportFooter(context, subjectCode.trim());
      });
      setFooterMsg({
        text: "Footer set from section 3 onward. Page numbers start at 3 on the Title Page.",
        isError: false,
      });
    } catch (err) {
      console.error(err);
      setFooterMsg({ text: err.message || "Could not set the footer.", isError: true });
    }
  }

  const msg = (m) => <Status text={m.text} tone={m.isError ? "error" : "success"} />;

  return (
    <StepBody>
      <Hint>
        Shown from the Title Page onward; the cover pages stay without a header or footer. Click
        again any time to change the text.
      </Hint>

      <Group>
        <SubHeading>Header</SubHeading>
        <Field label="Author's name (left)">
          <Input
            placeholder="e.g. Alkis A. Stavrinides"
            value={authorName}
            onChange={(_, data) => setAuthorName(data.value)}
          />
        </Field>
        <Field label="FYP title, shortened if needed (right)">
          <Input
            placeholder="e.g. Asset Management System"
            value={shortTitle}
            onChange={(_, data) => setShortTitle(data.value)}
          />
        </Field>
        <Actions>
          <Button appearance="primary" onClick={handleSetHeader}>
            Set Header
          </Button>
        </Actions>
        {msg(headerMsg)}
      </Group>

      <Group>
        <SubHeading>Footer</SubHeading>
        <Field label="Subject code (left)" hint='The right side shows "Page X of Y" automatically.'>
          <Input value={subjectCode} onChange={(_, data) => setSubjectCode(data.value)} />
        </Field>
        <Actions>
          <Button appearance="primary" onClick={handleSetFooter}>
            Set Footer
          </Button>
        </Actions>
        {msg(footerMsg)}
      </Group>
    </StepBody>
  );
}
