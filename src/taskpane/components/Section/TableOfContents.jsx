import React, { useState, useEffect, useCallback } from "react";
import { Button, Field, Input, Dropdown, Option } from "@fluentui/react-components";
import { ArrowSync20Regular, TextBulletListSquare20Regular, Add20Regular } from "@fluentui/react-icons";
import { StepBody, Hint, SubHeading, Group, Actions, Status, BusyBar, useBusy } from "../ui/StepUI";
import {
  generateTableOfContents,
  updateTableOfContents,
  listParentHeadings,
  addSubheading,
  refreshContentsPageNumbers,
} from "../Section/utils/toc-generator-utils";
import { REPORT_STRUCTURE } from "../Section/utils/toc-template-config";

/* global Word */

const LINK_TIP = "Ctrl+Click an entry to jump to that page.";

export default function TableOfContents() {
  const [genMsg, setGenMsg] = useState({ text: "", isError: false });
  const [updMsg, setUpdMsg] = useState({ text: "", isError: false });
  const [addMsg, setAddMsg] = useState({ text: "", isError: false });

  const [parents, setParents] = useState([]);
  const [parentPrefix, setParentPrefix] = useState("");
  const [title, setTitle] = useState("");

  const refreshParents = useCallback(async () => {
    try {
      await Word.run(async (context) => {
        const list = await listParentHeadings(context);
        setParents(list);
        setParentPrefix((current) => (list.some((p) => p.prefix === current) ? current : ""));
      });
    } catch (err) {
      console.error(err);
    }
  }, []);

  useEffect(() => {
    refreshParents();
  }, [refreshParents]);

  // Word finishes laying out the pages in the background AFTER the add-in's
  // changes; page numbers read before that all show the same early page.
  // So refresh them again a moment later (like pressing F9), twice to cover
  // slower/longer documents.
  function refreshPageNumbersLater() {
    [1500, 4000].forEach((delay) =>
      setTimeout(() => {
        Word.run((context) => refreshContentsPageNumbers(context)).catch((err) =>
          console.warn("Delayed page-number refresh failed (press F9 in Word).", err)
        );
      }, delay)
    );
  }

  async function run(setMsg, work, fallback) {
    try {
      await Word.run(async (context) => {
        setMsg({ text: await work(context), isError: false });
      });
      refreshPageNumbersLater();
    } catch (err) {
      console.error(err);
      // Word's errors are vague (e.g. "GeneralException"); add where it happened.
      const where = err?.debugInfo?.errorLocation ? ` (at ${err.debugInfo.errorLocation})` : "";
      setMsg({ text: `${err.message || fallback}${where}`, isError: true });
    }
  }

  async function handleGenerate() {
    await run(
      setGenMsg,
      async (context) => {
        const { created, count } = await generateTableOfContents(context, REPORT_STRUCTURE);
        return created
          ? `Contents page and all chapter headings generated (${count} entries). ${LINK_TIP}`
          : "Already generated. Use Update Table of Contents to refresh it.";
      },
      "Could not generate the Table of Contents."
    );
    refreshParents();
  }

  async function handleUpdate() {
    await run(
      setUpdMsg,
      async (context) => {
        const { count } = await updateTableOfContents(context);
        return `Table of Contents updated (${count} entries) with current page numbers. ${LINK_TIP}`;
      },
      "Could not update the Table of Contents."
    );
    refreshParents();
  }

  async function handleAdd() {
    if (!parentPrefix) {
      setAddMsg({ text: "Choose which heading to add it under.", isError: true });
      return;
    }
    await run(
      setAddMsg,
      async (context) => {
        const { text } = await addSubheading(context, parentPrefix, title);
        setTitle("");
        return `Added "${text}" and updated the Table of Contents.`;
      },
      "Could not add the subheading."
    );
    refreshParents();
  }

  // Loading bar: wraps the handlers above without changing them. The
  // "settle" phase covers the delayed page-number refresh (1.5 s and 4 s).
  const busy = useBusy();
  const SETTLE_MS = 4500;

  const selected = parents.find((p) => p.prefix === parentPrefix);
  const msg = (m) => <Status text={m.text} tone={m.isError ? "error" : "success"} />;

  return (
    <StepBody>
      <Group>
        <SubHeading>Create</SubHeading>
        <Hint>
          Builds the CONTENTS page and every chapter heading from the UPTM report template. Runs
          once - it won&apos;t overwrite anything you&apos;ve written.
        </Hint>
        <Actions stretch>
          <Button
            appearance="primary"
            icon={<TextBulletListSquare20Regular />}
            disabled={busy.active}
            onClick={busy.wrap("Generating the Table of Contents and chapters...", handleGenerate, {
              settle: SETTLE_MS,
              group: "create",
            })}
          >
            Generate Table of Contents
          </Button>
        </Actions>
        <BusyBar busy={busy} group="create" />
        {msg(genMsg)}
      </Group>

      <Group>
        <SubHeading>Keep it up to date</SubHeading>
        <Hint>
          Added, renamed or deleted a heading, or the pages moved? Update rebuilds the CONTENTS and
          refreshes every page number. Click it again after adding content.
        </Hint>
        <Actions stretch>
          <Button
            icon={<ArrowSync20Regular />}
            disabled={busy.active}
            onClick={busy.wrap("Updating the Table of Contents...", handleUpdate, {
              settle: SETTLE_MS,
              group: "update",
            })}
          >
            Update Table of Contents
          </Button>
        </Actions>
        <BusyBar busy={busy} group="update" />
        {msg(updMsg)}
      </Group>

      <Group>
        <SubHeading>Add a subheading</SubHeading>
        <Hint>
          Adds the next number under the heading you choose (under 3.3 it becomes 3.3.5), formats
          it, and updates the Table of Contents.
        </Hint>
        <Field label="Add under">
          <Dropdown
            placeholder="Choose a heading"
            value={selected ? selected.text : ""}
            selectedOptions={parentPrefix ? [parentPrefix] : []}
            onOptionSelect={(_, data) => setParentPrefix(data.optionValue)}
            onOpenChange={(_, data) => data.open && refreshParents()}
          >
            {parents.map((p) => (
              <Option key={p.prefix} value={p.prefix} text={p.text}>
                {`${"\u00a0\u00a0".repeat(p.level - 1)}${p.text}`}
              </Option>
            ))}
          </Dropdown>
        </Field>
        <Field label="Subheading title">
          <Input
            placeholder="e.g. Phase 5"
            value={title}
            onChange={(_, data) => setTitle(data.value)}
          />
        </Field>
        <Actions>
          <Button appearance="subtle" icon={<ArrowSync20Regular />} onClick={refreshParents}>
            Refresh list
          </Button>
          <Button
            appearance="primary"
            icon={<Add20Regular />}
            disabled={busy.active}
            onClick={busy.wrap("Adding the subheading...", handleAdd, { settle: SETTLE_MS, group: "add" })}
          >
            Add Subheading
          </Button>
        </Actions>
        <BusyBar busy={busy} group="add" />
        {msg(addMsg)}
      </Group>
    </StepBody>
  );
}
