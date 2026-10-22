import React, { useState, useEffect, useCallback } from "react";
import { Button, Field, Input, Dropdown, Option } from "@fluentui/react-components";
import {
  ArrowSync20Regular,
  TextBulletListSquare20Regular,
  Add20Regular,
  BookAdd20Regular,
} from "@fluentui/react-icons";
import { StepBody, SubHeading, Group, Actions, Status, BusyBar, useBusy, GuideSteps } from "../ui/StepUI";
import {
  generateTableOfContents,
  updateTableOfContents,
  listParentHeadings,
  addSubheading,
  addChapter,
  refreshAllPageNumbers,
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
  const [chapterTitle, setChapterTitle] = useState("");
  const [chapterMsg, setChapterMsg] = useState({ text: "", isError: false });

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
  // changes; page numbers read before that can be too early. So after every
  // Generate / Update / Add, CHECK every page number in the document (the
  // Contents and the Lists of Figures / Tables): refresh them all, wait,
  // refresh again, and stop once two checks in a row agree - then say so.
  function checkPageNumbersLater(setMsg, { retried = false } = {}) {
    // ms between checks: quick at first, then longer for big documents
    // (Word may still be laying out new pages) - about 20 s at most.
    const waits = [1200, 1300, 1500, 2000, 3000, 4000, 6000];
    let previous = null;
    const check = (i) =>
      setTimeout(async () => {
        let now;
        try {
          now = await Word.run((context) => refreshAllPageNumbers(context));
        } catch (err) {
          console.warn("Page-number check failed (press F9 in Word).", err);
          return;
        }
        const key = now.pages.join("|");
        // Done only when the numbers make sense AND two checks agree.
        const settled = now.ready && previous === key;
        previous = key;
        if (settled) {
          setMsg((m) =>
            m.isError || !now.pages.length
              ? m
              : { ...m, text: `${m.text} All ${now.pages.length} page numbers checked.` }
          );
          return;
        }
        if (i === waits.length - 1) {
          if (!retried) {
            // Still not right (e.g. just after Generate, Word was still laying
            // out the new pages): rebuild the Contents once more by itself -
            // the same as clicking Update - then check again.
            try {
              await Word.run((context) => updateTableOfContents(context));
            } catch (err) {
              console.warn("Automatic second update failed.", err);
            }
            checkPageNumbersLater(setMsg, { retried: true });
            return;
          }
          setMsg((m) =>
            m.isError
              ? m
              : {
                  ...m,
                  text: `${m.text} Some page numbers are not ready yet - scroll through your document once, then click Update Table of Contents.`,
                }
          );
          return;
        }
        check(i + 1);
      }, waits[i]);
    check(0);
  }

  // checkPages: false skips the page-number check (nothing in the
  // Contents changed, e.g. Add Subheading). `work` returns the message, or
  // { text, changed: false } when it changed nothing. Returns false when
  // nothing changed, so the loading bar skips its "checking" phase.
  async function run(setMsg, work, fallback, { checkPages = true } = {}) {
    try {
      let changed = true;
      await Word.run(async (context) => {
        const out = await work(context);
        const text = typeof out === "string" ? out : out.text;
        if (typeof out === "object" && out.changed === false) changed = false;
        setMsg({ text, isError: false });
      });
      if (checkPages && changed) checkPageNumbersLater(setMsg);
      return changed && checkPages;
    } catch (err) {
      console.error(err);
      // Word's errors are vague (e.g. "GeneralException"); add where it happened.
      const where = err?.debugInfo?.errorLocation ? ` (at ${err.debugInfo.errorLocation})` : "";
      setMsg({ text: `${err.message || fallback}${where}`, isError: true });
      return false;
    }
  }

  async function handleGenerate() {
    const result = await run(
      setGenMsg,
      async (context) => {
        const { created, count } = await generateTableOfContents(context, REPORT_STRUCTURE);
        return created
          ? `Contents page and all chapter headings generated (${count} entries). ${LINK_TIP}`
          : { text: "Already generated. Use Update Table of Contents to refresh it.", changed: false };
      },
      "Could not generate the Table of Contents."
    );
    refreshParents();
    return result;
  }

  async function handleUpdate() {
    const result = await run(
      setUpdMsg,
      async (context) => {
        const { count } = await updateTableOfContents(context);
        return `Table of Contents updated (${count} entries) with current page numbers. ${LINK_TIP}`;
      },
      "Could not update the Table of Contents."
    );
    refreshParents();
    return result;
  }

  async function handleAdd() {
    if (!parentPrefix) {
      setAddMsg({ text: "Choose which heading to add it under.", isError: true });
      return false;
    }
    const result = await run(
      setAddMsg,
      async (context) => {
        const { text } = await addSubheading(context, parentPrefix, title);
        setTitle("");
        return `Added "${text}". When you have finished adding subheadings, click Update Table of Contents (above) to list them with their page numbers.`;
      },
      "Could not add the subheading.",
      { checkPages: false }
    );
    refreshParents();
    return result;
  }

  async function handleAddChapter() {
    if (!chapterTitle.trim()) {
      setChapterMsg({ text: "Type a title for the new chapter.", isError: true });
      return false;
    }
    const result = await run(
      setChapterMsg,
      async (context) => {
        const { number, title: added } = await addChapter(context, chapterTitle);
        setChapterTitle("");
        return `Added CHAPTER ${number} - ${added}, with ${number}.1 INTRODUCTION, and updated the Table of Contents.`;
      },
      "Could not add the chapter."
    );
    refreshParents();
    return result;
  }

  // Loading bar: wraps the handlers above without changing them. The
  // "settle" phase covers the page-number checks after each action.
  const busy = useBusy();
  const SETTLE_MS = 8000; // the page-number checks keep going after this if needed

  const selected = parents.find((p) => p.prefix === parentPrefix);
  const msg = (m) => <Status text={m.text} tone={m.isError ? "error" : "success"} />;

  return (
    <StepBody>
      <GuideSteps
        steps={[
          <>Click <b>Generate Table of Contents</b> (only once). It adds the Contents page and all chapter headings.</>,
          <>Write your report under each heading in Word.</>,
          <>Need another chapter? Type its title in <b>Add a chapter</b> and click <b>Add Chapter</b>.</>,
          <>Need another heading? In <b>Add a subheading</b>, choose where it goes, type the title and click <b>Add Subheading</b>. Under a 3-digit heading (e.g. 2.3.4) it is lettered a., b., c.</>,
          <>When you finish adding headings, click <b>Update Table of Contents</b> to list them and fix the page numbers.</>,
        ]}
        note="Only chapters and headings up to 3 digits (e.g. 2.3.4) are listed. Ctrl+Click a line in the Contents to jump to it."
      />
      <Group>
        <SubHeading>Create</SubHeading>
        <Actions stretch>
          <Button
            appearance="primary"
            icon={<TextBulletListSquare20Regular />}
            disabled={busy.active}
            onClick={busy.wrap("Generating the Table of Contents and chapters...", handleGenerate, {
              settle: SETTLE_MS,
              settleLabel: "Checking every page number...",
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
        <Actions stretch>
          <Button
            icon={<ArrowSync20Regular />}
            disabled={busy.active}
            onClick={busy.wrap("Updating the Table of Contents...", handleUpdate, {
              settle: SETTLE_MS,
              settleLabel: "Checking every page number...",
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
        <SubHeading>Add a chapter</SubHeading>
        <Field label="Chapter title">
          <Input
            placeholder="e.g. Testing and Evaluation"
            value={chapterTitle}
            onChange={(_, data) => setChapterTitle(data.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter" && !busy.active) {
                busy.wrap("Adding the chapter...", handleAddChapter, { settle: SETTLE_MS, settleLabel: "Checking every page number...", group: "chapter" })();
              }
            }}
          />
        </Field>
        <Actions>
          <Button
            appearance="primary"
            icon={<BookAdd20Regular />}
            disabled={busy.active}
            onClick={busy.wrap("Adding the chapter...", handleAddChapter, { settle: SETTLE_MS, settleLabel: "Checking every page number...", group: "chapter" })}
          >
            Add Chapter
          </Button>
        </Actions>
        <BusyBar busy={busy} group="chapter" />
        {msg(chapterMsg)}
      </Group>

      <Group>
        <SubHeading>Add a subheading</SubHeading>
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
            onClick={busy.wrap("Adding the subheading...", handleAdd, { group: "add" })}
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
