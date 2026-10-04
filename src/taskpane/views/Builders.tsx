import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Button, Field, Select, Text, Textarea } from "@fluentui/react-components";
import { activity, describeError } from "../../lib/activity";
import { parseKpis, type KpiSource, type KpiStyle } from "../../lib/kpi";
import { insertKpis, readKpiSelection } from "../../lib/kpiActions";
import { UserError } from "../../lib/ppt";
import { prefs } from "../../lib/prefs";
import { parseRoadmap, quarterLabel, type RoadmapSource } from "../../lib/roadmap";
import { insertRoadmap, readRoadmapSelection } from "../../lib/roadmapActions";
import { pane } from "../store";
import { Section, Seg, useUi } from "../ui";

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

async function run(work: () => Promise<string>) {
  try {
    pane.showMessage({ intent: "success", text: await activity.track("pane builder", work) });
  } catch (e) {
    pane.showMessage({ intent: e instanceof UserError ? "warning" : "error", text: e instanceof UserError ? e.message : describeError(e) });
  }
}

/**
 * Follow the selection: when a roadmap or KPI shapes Retro built are selected, hand back their
 * source so the builder opens them for editing.
 */
function useBuiltSelection<T>(read: () => Promise<{ id: string; source: T } | undefined>, open: (source: T) => void) {
  const [id, setId] = useState<string | undefined>();
  const last = useRef<string | undefined>(undefined);
  const check = useCallback(() => {
    if (typeof PowerPoint === "undefined") return;
    read()
      .then((found) => {
        // Open it only when a different one is selected, so edits in progress aren't replaced.
        if (found && found.id !== last.current) open(found.source);
        last.current = found?.id;
        setId(found?.id);
      })
      .catch(() => setId(undefined));
  }, [read, open]);
  useEffect(() => {
    check();
    const doc = typeof Office === "undefined" ? undefined : Office.context?.document;
    if (!doc?.addHandlerAsync) return;
    doc.addHandlerAsync(Office.EventType.DocumentSelectionChanged, check);
    return () => doc.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: check });
  }, [check]);
  return id;
}

function RoadmapBuilder() {
  const ui = useUi();
  const [text, setText] = useState("");
  const [fyStart, setFyStart] = useState(1);
  useEffect(() => {
    void prefs.fyStart().then(setFyStart);
  }, []);
  const open = useCallback((source: RoadmapSource) => {
    setText(source.text);
    setFyStart(source.fyStart);
  }, []);
  const selected = useBuiltSelection(readRoadmapSelection, open);
  const parsed = useMemo(() => parseRoadmap(text, fyStart), [text, fyStart]);
  const quarters = parsed.items.flatMap((i) => [i.start, i.end ?? i.start]);
  const lanes = new Set(parsed.items.map((i) => i.lane)).size;

  return (
    <Section title="Roadmap">
      <Field label="Fiscal year starts in">
        <Select
          value={String(fyStart)}
          onChange={(_, d) => {
            setFyStart(Number(d.value));
            void prefs.setFyStart(Number(d.value));
          }}
        >
          {MONTHS.map((m, i) => (
            <option key={m} value={i + 1}>
              {m}
            </option>
          ))}
        </Select>
      </Field>
      <Field label="One row per item: lane, item, start, end" hint="Paste from Excel, or separate columns with | or commas. Leave the end empty for a milestone. Dates: Q1 FY27, Q1 2027 or Mar 2027.">
        <Textarea value={text} onChange={(_, d) => setText(d.value)} placeholder={"Platform | New API | Q1 FY27 | Q2 FY27\nPlatform | GA launch | Q3 FY27\nData | Lakehouse | Q2 FY27 | Q4 FY27"} rows={6} resize="vertical" />
      </Field>
      {text.trim() && (
        <Text size={200} className={ui.muted}>
          {parsed.items.length
            ? `${parsed.items.length} item${parsed.items.length === 1 ? "" : "s"} in ${lanes} lane${lanes === 1 ? "" : "s"}, ${quarterLabel(Math.min(...quarters), fyStart)} to ${quarterLabel(Math.max(...quarters), fyStart)}.`
            : "No items yet."}
          {parsed.errors.map((e) => ` ${e}`).join("")}
        </Text>
      )}
      <Button appearance="primary" style={{ alignSelf: "flex-start" }} disabled={!parsed.items.length} onClick={() => void run(() => insertRoadmap({ text, fyStart }, selected))}>
        {selected ? "Update roadmap" : "Draw roadmap"}
      </Button>
      <Text size={200} className={ui.muted}>
        {selected
          ? "Editing the selected roadmap: changes replace it in the same spot."
          : "Draws below the slide title, in your brand colors and body font, as one group you can edit. Select it later to change the rows."}
      </Text>
    </Section>
  );
}

function KpiBuilder() {
  const ui = useUi();
  const [text, setText] = useState("");
  const [style, setStyle] = useState<KpiStyle>("tiles");
  const open = useCallback((source: KpiSource) => {
    setText(source.text);
    setStyle(source.style);
  }, []);
  const selected = useBuiltSelection(readKpiSelection, open);
  const parsed = useMemo(() => parseKpis(text), [text]);

  return (
    <Section title="KPIs and scorecard">
      <Seg
        label="Make"
        value={style}
        options={[
          { value: "tiles", label: "Big-number tiles" },
          { value: "table", label: "Scorecard table" },
        ]}
        onChange={setStyle}
      />
      <Field label="One row per metric: metric, value, change, status" hint="Paste from Excel, or separate columns with | or commas. Change like +12% or −3 pts; status G, A or R (or green, at risk, off track…). Both optional.">
        <Textarea value={text} onChange={(_, d) => setText(d.value)} placeholder={"ARR | $4.2M | +12% | G\nChurn | 2.1% | -0.3 pts | A\nNPS | 48 | +4 | G"} rows={5} resize="vertical" />
      </Field>
      {text.trim() && (
        <Text size={200} className={ui.muted}>
          {parsed.rows.length ? `${parsed.rows.length} metric${parsed.rows.length === 1 ? "" : "s"}.` : "No metrics yet."}
          {parsed.errors.map((e) => ` ${e}`).join("")}
        </Text>
      )}
      <Button appearance="primary" style={{ alignSelf: "flex-start" }} disabled={!parsed.rows.length} onClick={() => void run(() => insertKpis({ text, style }, selected))}>
        {selected ? "Update" : style === "tiles" ? "Make tiles" : "Make table"}
      </Button>
      <Text size={200} className={ui.muted}>
        {selected
          ? "Editing the selected KPIs: changes replace them in the same spot."
          : "Values use your brand's first color; changes and status show green, amber or red. Select them later to change the rows."}
      </Text>
    </Section>
  );
}

export function Builders() {
  return (
    <>
      <RoadmapBuilder />
      <KpiBuilder />
    </>
  );
}
