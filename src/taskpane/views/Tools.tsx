import { useCallback, useEffect, useState } from "react";
import { Button, Field, Text, Textarea, makeStyles, tokens } from "@fluentui/react-components";
import { EraserRegular } from "@fluentui/react-icons";
import { addStamp, addSticky, scanNotes, STAMPS, STICKY_COLORS, type NoteCount, type StampLabel, type StickyColor } from "../../lib/notes";
import { activity, describeError } from "../../lib/activity";
import { selectedShapes, UserError } from "../../lib/ppt";
import { pane, usePane } from "../store";
import { useSmartSelection } from "../useSmartSelection";
import { SmartEditor, SmartInsert } from "./SmartSection";
import { AppendixSection } from "./AppendixSection";
import { Builders } from "./Builders";
import { TrackerSection } from "./TrackerSection";
import { Section, Seg, useUi } from "../ui";

const useStyles = makeStyles({
  swatches: { display: "flex", gap: "8px" },
  swatch: {
    width: "28px",
    height: "28px",
    borderRadius: "4px",
    border: "1px solid rgba(0,0,0,.18)",
    cursor: "pointer",
    padding: 0,
    ":focus-visible": { outline: `2px solid ${tokens.colorStrokeFocus2}`, outlineOffset: "2px" },
  },
  on: { boxShadow: `0 0 0 2px ${tokens.colorNeutralBackground1}, 0 0 0 4px ${tokens.colorNeutralForeground1}` },
  between: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: "8px" },
});

const COLOR_NAMES: Record<StickyColor, string> = { yellow: "Yellow", pink: "Pink", blue: "Blue", green: "Green" };

async function report(work: () => Promise<string | void>) {
  try {
    const text = await activity.track("pane tools", work);
    pane.showMessage(text ? { intent: "success", text } : undefined);
  } catch (e) {
    pane.showMessage({ intent: e instanceof UserError ? "warning" : "error", text: e instanceof UserError ? e.message : describeError(e) });
  }
}

export function Tools() {
  const s = useStyles();
  const ui = useUi();
  const [text, setText] = useState("");
  const [color, setColor] = useState<StickyColor>("yellow");
  const [stamp, setStamp] = useState<StampLabel>("DRAFT");
  const [count, setCount] = useState<NoteCount | undefined>();

  const recount = useCallback(() => {
    if (typeof PowerPoint === "undefined") return;
    scanNotes(false)
      .then(setCount)
      .catch(() => setCount(undefined));
  }, []);

  useEffect(() => {
    // Count existing notes and stamps when the section opens.
    recount();
  }, [recount]);

  const plural = (n: number, word: string) => `${n} ${word}${n === 1 ? "" : "s"}`;

  const smart = useSmartSelection();

  const { toolsMode: mode } = usePane();

  return (
    <>
      <Section>
        <Seg
          label="Tools"
          value={mode}
          options={[
            { value: "notes", label: "Notes" },
            { value: "smart", label: "Smart" },
            { value: "sections", label: "Sections" },
            { value: "builders", label: "Builders" },
          ]}
          onChange={pane.setToolsMode}
        />
      </Section>
      {smart.selected && <SmartEditor selected={smart.selected} refresh={smart.refresh} />}
      {mode === "notes" && (
        <>
          <Section title="Sticky note">
            <Field label="Note">
              <Textarea value={text} placeholder="Check this number with finance" onChange={(_, d) => setText(d.value)} resize="vertical" />
            </Field>
            <div className={s.between}>
              <div className={s.swatches} role="radiogroup" aria-label="Note color">
                {(Object.keys(STICKY_COLORS) as StickyColor[]).map((c) => (
                  <button
                    key={c}
                    className={`${s.swatch} ${c === color ? s.on : ""}`}
                    style={{ background: STICKY_COLORS[c] }}
                    role="radio"
                    aria-checked={c === color}
                    aria-label={COLOR_NAMES[c]}
                    onClick={() => setColor(c)}
                  />
                ))}
              </div>
              <Button
                appearance="primary"
                onClick={() =>
                  void report(async () => {
                    const selection = await PowerPoint.run(async (context) => (await selectedShapes(context)).rects);
                    await addSticky(text.trim() || "Note", color, selection);
                    recount();
                  })
                }
              >
                Add sticky
              </Button>
            </div>
          </Section>

          <Section title="Stamp">
            <Seg label="Stamp text" value={stamp} options={STAMPS.map((v) => ({ value: v, label: v }))} onChange={setStamp} />
            <div className={ui.grid2}>
              <Button
                onClick={() =>
                  void report(async () => {
                    await addStamp(stamp);
                    recount();
                  })
                }
              >
                This slide
              </Button>
              <Button
                onClick={() =>
                  void report(async () => {
                    const n = await addStamp(stamp, true);
                    recount();
                    return `Stamped ${plural(n, "slide")}.`;
                  })
                }
              >
                All slides
              </Button>
            </div>
          </Section>

          <Section>
            <div className={s.between}>
              <div>
                <b>{count ? plural(count.notes, "note") + " & stamps" : "Notes & stamps"}</b>
                <br />
                <Text size={200} className={ui.muted}>
                  {count ? `on ${plural(count.slides, "slide")}` : "Counting…"}
                </Text>
              </div>
              <Button
                icon={<EraserRegular />}
                disabled={!count?.notes}
                onClick={() =>
                  void report(async () => {
                    const r = await scanNotes(true);
                    recount();
                    return `Removed ${plural(r.notes, "note")} and stamps from ${plural(r.slides, "slide")}.`;
                  })
                }
              >
                Remove all
              </Button>
            </div>
            <Text size={200} className={ui.muted}>
              Retro only removes the notes and stamps it added, so the rest of your slides are never touched.
            </Text>
          </Section>
        </>
      )}

      {mode === "smart" && <SmartInsert hasSelection={Boolean(smart.selected)} refresh={smart.refresh} />}

      {mode === "sections" && (
        <>
          <TrackerSection />
          <AppendixSection />
        </>
      )}

      {mode === "builders" && <Builders />}
    </>
  );
}
