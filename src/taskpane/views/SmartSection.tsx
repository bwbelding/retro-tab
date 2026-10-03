import { useEffect, useState } from "react";
import { Button, Field, Input, Text } from "@fluentui/react-components";
import { UserError } from "../../lib/ppt";
import { insertSmart, renumberSlide, SMART, SMART_KINDS, updateSmart, type SmartKind } from "../../lib/smart";
import { pane } from "../store";
import { Section, Seg, useUi } from "../ui";
import { useSmartSelection } from "../useSmartSelection";

async function report(work: () => Promise<void>) {
  try {
    await work();
    pane.showMessage(undefined);
  } catch (e) {
    pane.showMessage({ intent: e instanceof UserError ? "warning" : "error", text: e instanceof Error ? e.message : String(e) });
  }
}

export function SmartSection() {
  const ui = useUi();
  const { selected, refresh } = useSmartSelection();
  const [typed, setTyped] = useState("");

  useEffect(() => {
    // Show the selected element's current value in the number field.
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setTyped(selected?.value ?? "");
  }, [selected]);

  const set = (value: string) =>
    selected &&
    report(async () => {
      await updateSmart(selected, value);
      refresh();
    });

  const insert = (kind: SmartKind) =>
    report(async () => {
      await insertSmart(kind);
      refresh();
    });

  const info = selected && SMART[selected.kind];

  return (
    <Section title="Smart elements">
      <div className={ui.grid2}>
        {SMART_KINDS.map((k) => (
          <Button key={k} size="small" onClick={() => void insert(k)}>
            {SMART[k].label}
          </Button>
        ))}
      </div>

      {selected && info ? (
        <>
          <Text weight="semibold">Selected: {info.label}</Text>
          {info.choices ? (
            <Seg label={`${info.label} value`} value={selected.value} options={info.choices} onChange={(v) => void set(v)} />
          ) : (
            <>
              <div className={ui.row}>
                <Field label={selected.kind === "progress" ? "Percent" : "Number"} style={{ flex: 1 }}>
                  <Input type="number" value={typed} onChange={(_, d) => setTyped(d.value)} />
                </Field>
                <Button appearance="primary" onClick={() => void set(typed)}>
                  Set
                </Button>
              </div>
              {selected.kind === "progress" && (
                <Seg
                  label="Quick percent"
                  value={selected.value}
                  options={["0", "25", "50", "75", "100"].map((v) => ({ value: v, label: `${v}%` }))}
                  onChange={(v) => void set(v)}
                />
              )}
              {selected.kind === "number" && (
                <Button
                  onClick={() =>
                    void report(async () => {
                      await renumberSlide();
                      refresh();
                    })
                  }
                >
                  Renumber this slide 1, 2, 3…
                </Button>
              )}
            </>
          )}
        </>
      ) : (
        <Text size={200} className={ui.muted}>
          Select a Smart Element on the slide to change it. Click the element's edge so the whole element is selected.
        </Text>
      )}
    </Section>
  );
}
