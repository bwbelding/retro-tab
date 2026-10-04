import { useEffect, useState } from "react";
import { Button, Checkbox, Field, Input, Select, Text, makeStyles, tokens } from "@fluentui/react-components";
import { ArrowDownRegular, ArrowLeftRegular, ArrowRightRegular, ArrowSwapRegular, ArrowUpRegular } from "@fluentui/react-icons";
import {
  commonValue,
  distributeWithGap,
  fromPoints,
  matchSize,
  nudge,
  setAsWhole,
  setEach,
  swapPositions,
  toPoints,
  type Axis,
  type Dimension,
  type ExactValues,
  type Rect,
  type Reference,
  type Unit,
} from "../../lib/layout";
import { arrangeLogos } from "../../lib/logoGridActions";
import { applyRects, requireSelection, UserError } from "../../lib/ppt";
import { prefs } from "../../lib/prefs";
import { pane } from "../store";
import { Section, Seg, useUi } from "../ui";
import { useSelection } from "../useSelection";

const useStyles = makeStyles({
  status: {
    padding: "8px 16px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    fontSize: tokens.fontSizeBase200,
  },
  input: { minWidth: 0, width: "100%" },
  pad: { display: "grid", gridTemplateColumns: "repeat(3, 32px)", gridTemplateRows: "repeat(3, 32px)", gap: "2px" },
});

const UNITS = [
  { value: "pt", label: "pt" },
  { value: "cm", label: "cm" },
  { value: "in", label: "in" },
] as const;

const REFERENCES = [
  { value: "first", label: "First" },
  { value: "last", label: "Last" },
  { value: "largest", label: "Largest" },
  { value: "smallest", label: "Smallest" },
] as const;

const round = (n: number) => String(Math.round(n * 100) / 100);
const parse = (s: string) => (s.trim() === "" ? undefined : Number(s));

/** Read the selection, transform its rectangles, and write them back. Problems show in the pane. */
async function transform(min: number, what: string, fn: (rects: Rect[]) => Rect[]) {
  try {
    await PowerPoint.run(async (context) => {
      const { shapes, rects } = await requireSelection(context, min, what);
      applyRects(shapes, fn(rects));
      await context.sync();
    });
    pane.showMessage(undefined);
  } catch (e) {
    pane.showMessage({
      intent: e instanceof UserError ? "warning" : "error",
      text: e instanceof Error ? e.message : String(e),
    });
  }
}

export function Layout() {
  const s = useStyles();
  const ui = useUi();
  const { rects, refresh } = useSelection();
  const [unit, setUnit] = useState<Unit>("pt");
  const [ref, setRef] = useState<Reference>("first");
  const [gap, setGap] = useState("12");
  const [step, setStep] = useState("1");
  const [exact, setExact] = useState({ left: "", top: "", width: "", height: "" });
  const [lockAspect, setLockAspect] = useState(false);
  const [whole, setWhole] = useState<"each" | "whole">("each");

  useEffect(() => {
    // Load the saved reference choice once.
    void prefs.matchReference().then(setRef);
  }, []);

  useEffect(() => {
    // Show the selection's shared values (blank when shapes differ) in the current unit.
    const show = (k: keyof Omit<Rect, "id">) => {
      const v = commonValue(rects, k);
      return v === undefined ? "" : round(fromPoints(v, unit));
    };
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setExact({ left: show("left"), top: show("top"), width: show("width"), height: show("height") });
  }, [rects, unit]);

  const pts = (v: string | undefined) => (v === undefined || Number.isNaN(Number(v)) ? undefined : toPoints(Number(v), unit));
  const n = rects.length;
  const done = () => refresh();

  const doMatch = (dim: Dimension) => transform(2, "match their sizes", (r) => matchSize(r, dim, ref)).then(done);
  const doGap = (axis: Axis) => {
    const g = pts(gap);
    if (g === undefined) return pane.showMessage({ intent: "warning", text: "Enter a gap, for example 12." });
    return transform(2, "space them", (r) => distributeWithGap(r, g, axis)).then(done);
  };
  const doNudge = (dx: number, dy: number) => {
    const d = pts(step) ?? 0;
    return transform(1, "nudge them", (r) => nudge(r, dx * d, dy * d)).then(done);
  };
  const doExact = () => {
    const v: ExactValues = {};
    for (const k of ["left", "top", "width", "height"] as const) {
      const p = parse(exact[k]);
      if (p !== undefined) {
        if (Number.isNaN(p)) return pane.showMessage({ intent: "warning", text: "Use numbers only in X, Y, Width and Height." });
        v[k] = toPoints(p, unit);
      }
    }
    return transform(1, "set their size and position", (r) => (whole === "whole" ? setAsWhole(r, v, lockAspect) : setEach(r, v, lockAspect))).then(done);
  };

  return (
    <>
      <div className={s.status} role="status">
        <b>{n === 0 ? "Nothing selected" : `${n} shape${n === 1 ? "" : "s"} selected`}</b>
      </div>

      <Section title="Units">
        <Seg label="Units" value={unit} options={UNITS} onChange={setUnit} />
      </Section>

      <Section title="Match size">
        <Field label="Match to">
          <Seg
            label="Match to"
            value={ref}
            options={REFERENCES}
            onChange={(r) => {
              setRef(r);
              void prefs.setMatchReference(r);
            }}
          />
        </Field>
        <Text size={200} className={ui.muted}>
          The ribbon's Match buttons use this choice too.
        </Text>
        <div className={ui.grid3}>
          <Button disabled={n < 2} onClick={() => void doMatch("width")}>Width</Button>
          <Button disabled={n < 2} onClick={() => void doMatch("height")}>Height</Button>
          <Button disabled={n < 2} onClick={() => void doMatch("both")}>Both</Button>
        </div>
      </Section>

      <Section title="Swap positions">
        <div className={ui.row} style={{ alignItems: "center" }}>
          <Button
            icon={<ArrowSwapRegular />}
            disabled={n !== 2}
            onClick={() => void transform(2, "swap their positions", (r) => [...swapPositions(r[0], r[1])]).then(done)}
          >
            Swap
          </Button>
          {n !== 2 && (
            <Text size={200} className={ui.muted}>
              Select exactly 2 shapes
            </Text>
          )}
        </div>
      </Section>

      <Section title="Distribute with a fixed gap">
        <Field label={`Gap (${unit})`}>
          <Input className={s.input} type="number" value={gap} onChange={(_, d) => setGap(d.value)} />
        </Field>
        <div className={ui.grid2}>
          <Button disabled={n < 2} onClick={() => void doGap("horizontal")}>Horizontally</Button>
          <Button disabled={n < 2} onClick={() => void doGap("vertical")}>Vertically</Button>
        </div>
      </Section>

      <Section title="Nudge">
        <div className={ui.row} style={{ alignItems: "center", gap: 16 }}>
          <div className={s.pad}>
            <span />
            <Button aria-label="Nudge up" icon={<ArrowUpRegular />} disabled={n < 1} onClick={() => void doNudge(0, -1)} />
            <span />
            <Button aria-label="Nudge left" icon={<ArrowLeftRegular />} disabled={n < 1} onClick={() => void doNudge(-1, 0)} />
            <span />
            <Button aria-label="Nudge right" icon={<ArrowRightRegular />} disabled={n < 1} onClick={() => void doNudge(1, 0)} />
            <span />
            <Button aria-label="Nudge down" icon={<ArrowDownRegular />} disabled={n < 1} onClick={() => void doNudge(0, 1)} />
            <span />
          </div>
          <Field label={`Step (${unit})`} style={{ flex: 1 }}>
            <Input className={s.input} type="number" value={step} onChange={(_, d) => setStep(d.value)} />
          </Field>
        </div>
      </Section>

      <Section title="Exact size & position">
        <div className={ui.grid2}>
          {(
            [
              ["left", "X"],
              ["top", "Y"],
              ["width", "Width"],
              ["height", "Height"],
            ] as const
          ).map(([k, label]) => (
            <Field key={k} label={`${label} (${unit})`}>
              <Input
                className={s.input}
                type="number"
                placeholder={n > 1 ? "Mixed" : ""}
                value={exact[k]}
                onChange={(_, d) => setExact((e) => ({ ...e, [k]: d.value }))}
              />
            </Field>
          ))}
        </div>
        <Checkbox label="Lock aspect ratio" checked={lockAspect} onChange={(_, d) => setLockAspect(Boolean(d.checked))} />
        <Field label="Apply to">
          <Seg
            label="Apply to"
            value={whole}
            options={[
              { value: "each", label: "Each shape" },
              { value: "whole", label: "Selection as a whole" },
            ]}
            onChange={setWhole}
          />
        </Field>
        <Button appearance="primary" disabled={n < 1} style={{ alignSelf: "flex-start" }} onClick={() => void doExact()}>
          Apply
        </Button>
      </Section>

      <LogoGridSection selected={n} />
    </>
  );
}

/** Arrange the selected logos in an even grid below the title, each equally prominent. */
function LogoGridSection({ selected }: { selected: number }) {
  const ui = useUi();
  const [columns, setColumns] = useState(0);
  return (
    <Section title="Logo grid">
      <Field label="Columns">
        <Select value={String(columns)} onChange={(_, d) => setColumns(Number(d.value))}>
          <option value="0">Automatic</option>
          {[2, 3, 4, 5, 6].map((c) => (
            <option key={c} value={c}>
              {c}
            </option>
          ))}
        </Select>
      </Field>
      <Button
        appearance="primary"
        disabled={selected < 2}
        style={{ alignSelf: "flex-start" }}
        onClick={() =>
          void arrangeLogos(columns || undefined)
            .then((count) => pane.showMessage({ intent: "success", text: `Arranged ${count} logos in a grid below the title.` }))
            .catch((e) => pane.showMessage({ intent: e instanceof UserError ? "warning" : "error", text: e instanceof Error ? e.message : String(e) }))
        }
      >
        Arrange selected logos
      </Button>
      <Text size={200} className={ui.muted}>
        Select the logos (any pictures), then arrange them: same spacing, each sized to look equally prominent whatever its shape. To insert logos from your photo folder as a grid, use Photos → Logo grid.
      </Text>
    </Section>
  );
}
