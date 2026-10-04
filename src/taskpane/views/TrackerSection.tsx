import { useCallback, useEffect, useState } from "react";
import { Button, Checkbox, Field, Input, Text, makeStyles, mergeClasses, tokens } from "@fluentui/react-components";
import { AddRegular, DeleteRegular, EditRegular } from "@fluentui/react-icons";
import { activity, describeError } from "../../lib/activity";
import { UserError } from "../../lib/ppt";
import { sectionOf, slideRange, type TrackerPosition, type TrackerSettings, type TrackerStyle } from "../../lib/tracker";
import { applyTracker, changeSection, goToSlide, markSection, readTracker, removeTracker, type TrackerState } from "../../lib/trackerActions";
import { pane } from "../store";
import { Section, Seg, useUi } from "../ui";

const useStyles = makeStyles({
  preview: { display: "flex", gap: "2px" },
  tab: {
    flex: "1 1 0",
    minWidth: 0,
    padding: "3px 4px",
    fontSize: tokens.fontSizeBase100,
    textAlign: "center",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground3,
    borderRadius: tokens.borderRadiusSmall,
  },
  tabOn: { backgroundColor: tokens.colorBrandBackground, color: tokens.colorNeutralForegroundOnBrand, fontWeight: tokens.fontWeightSemibold },
  table: { display: "grid", gridTemplateColumns: "minmax(0, 1fr) auto auto", gap: "4px 8px", alignItems: "center" },
  head: { fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3 },
  name: { overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", minWidth: 0 },
  current: { fontWeight: tokens.fontWeightSemibold },
  actions: { display: "flex", gap: "4px" },
  form: { display: "flex", flexDirection: "column", gap: "8px" },
});

const STYLES: { value: TrackerStyle; label: string }[] = [
  { value: "tabs", label: "Tabs" },
  { value: "bar", label: "Bar" },
  { value: "dots", label: "Dots" },
];
const POSITIONS: { value: TrackerPosition; label: string }[] = [
  { value: "top", label: "Top" },
  { value: "bottom", label: "Bottom" },
];

async function run(work: () => Promise<string | void>) {
  try {
    const text = await activity.track("pane tracker", work);
    pane.showMessage(text ? { intent: "success", text } : undefined);
  } catch (e) {
    pane.showMessage({ intent: e instanceof UserError ? "warning" : "error", text: e instanceof UserError ? e.message : describeError(e) });
  }
}

export function TrackerSection() {
  const s = useStyles();
  const ui = useUi();
  const [state, setState] = useState<TrackerState | undefined>();
  const [settings, setSettings] = useState<TrackerSettings | undefined>();
  // Naming a new section ("new"), renaming one (its start slide), or neither.
  const [naming, setNaming] = useState<"new" | number | undefined>();
  const [name, setName] = useState("");
  const [busy, setBusy] = useState(false);

  const reload = useCallback(() => {
    if (typeof PowerPoint === "undefined") return;
    readTracker()
      .then((next) => {
        setState(next);
        setSettings((cur) => cur ?? next.settings);
      })
      .catch(() => undefined);
  }, []);

  useEffect(() => {
    // Follow the selected slide, so the preview and "Add section at slide N" stay current.
    reload();
    const doc = typeof Office === "undefined" ? undefined : Office.context?.document;
    if (!doc?.addHandlerAsync) return;
    doc.addHandlerAsync(Office.EventType.DocumentSelectionChanged, reload);
    return () => doc.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: reload });
  }, [reload]);

  if (!state || !settings) {
    return (
      <Section title="Section tracker">
        <Text size={200} className={ui.muted}>
          Loading…
        </Text>
      </Section>
    );
  }

  const act = async (work: () => Promise<string | void>) => {
    setBusy(true);
    await run(work);
    setBusy(false);
    reload();
  };
  const active = sectionOf(state.sections, state.current);
  const startsHere = state.sections.some((x) => x.start === state.current);
  const submit = () => {
    const target = naming;
    setNaming(undefined);
    void act(() => (target === "new" ? markSection(name) : changeSection(target!, name)));
  };

  return (
    <Section title="Section tracker">
      {state.sections.length > 0 && (
        <>
          <div className={s.preview} aria-hidden>
            {state.sections.map((x, i) => (
              <div key={x.start} className={mergeClasses(s.tab, i === active && s.tabOn)}>
                {x.name}
              </div>
            ))}
          </div>
          <Text size={200} className={ui.muted}>
            {state.current < 0
              ? "Select a slide to preview its tracker."
              : state.current === state.agenda
                ? `Slide ${state.current + 1} is the agenda slide: no tracker.`
                : active < 0
                  ? `Slide ${state.current + 1} is before the first section: no tracker.`
                  : `Preview · slide ${state.current + 1}`}
          </Text>
          <div className={s.table}>
            <span className={s.head}>Section</span>
            <span className={s.head}>Slides</span>
            <span />
            {state.sections.map((x, i) => (
              <div key={x.start} style={{ display: "contents" }}>
                <Button size="small" appearance="outline" className={mergeClasses(s.name, i === active && s.current)} title={`Go to slide ${x.start + 1}`} onClick={() => void goToSlide(x.start)}>
                  {x.name}
                </Button>
                <Text size={200}>{slideRange(x)}</Text>
                <div className={s.actions}>
                  <Button
                    size="small"
                    appearance="outline"
                    icon={<EditRegular />}
                    aria-label={`Rename ${x.name}`}
                    onClick={() => {
                      setName(x.name);
                      setNaming(x.start);
                    }}
                  />
                  <Button size="small" appearance="outline" icon={<DeleteRegular />} aria-label={`Remove ${x.name}`} disabled={busy} onClick={() => void act(() => changeSection(x.start, null))} />
                </div>
              </div>
            ))}
          </div>
        </>
      )}

      {naming !== undefined ? (
        <form
          className={s.form}
          onSubmit={(e) => {
            e.preventDefault();
            submit();
          }}
        >
          <Field label={naming === "new" ? `Section starting at slide ${state.current + 1}` : "Section name"}>
            <Input value={name} onChange={(_, d) => setName(d.value)} placeholder="e.g. Market" autoFocus />
          </Field>
          <div className={ui.row}>
            <Button appearance="primary" type="submit">
              {naming === "new" ? "Add" : "Save"}
            </Button>
            <Button onClick={() => setNaming(undefined)}>Cancel</Button>
          </div>
        </form>
      ) : (
        <Button
          icon={<AddRegular />}
          style={{ alignSelf: "flex-start" }}
          disabled={state.current < 0 || startsHere || busy}
          title={startsHere ? "A section already starts at this slide. Use its rename button." : undefined}
          onClick={() => {
            setName("");
            setNaming("new");
          }}
        >
          {state.current < 0 ? "Add section" : `Add section at slide ${state.current + 1}`}
        </Button>
      )}
      {state.sections.length === 0 && (
        <Text size={200} className={ui.muted}>
          Select the first slide of a section and click Add section. Each section runs until the next one starts, so moving or adding slides keeps them right.
        </Text>
      )}

      <Field label="Style">
        <Seg label="Style" value={settings.style} options={STYLES} onChange={(style) => setSettings({ ...settings, style })} />
      </Field>
      <Field label="Position">
        <Seg label="Position" value={settings.position} options={POSITIONS} onChange={(position) => setSettings({ ...settings, position })} />
      </Field>
      <Checkbox label="Also build an agenda slide" checked={settings.agenda} onChange={(_, d) => setSettings({ ...settings, agenda: Boolean(d.checked) })} />
      <div className={ui.grid2}>
        <Button appearance="primary" disabled={busy || state.sections.length === 0} onClick={() => void act(() => applyTracker(settings))}>
          Apply tracker
        </Button>
        <Button disabled={busy} onClick={() => void act(removeTracker)}>
          Remove tracker
        </Button>
      </div>
      <Text size={200} className={ui.muted}>
        Uses your brand kit's first color and body font. After adding, moving or removing slides, click Apply again or use Refresh tracker on the ribbon.
      </Text>
    </Section>
  );
}
