import { useEffect, useState } from "react";
import { Button, Field, Input, Select, Text, makeStyles, tokens } from "@fluentui/react-components";
import { AddRegular, ArrowLeftRegular, ArrowRightRegular, ColorRegular, DeleteRegular, DismissRegular, DocumentArrowDownRegular, EditRegular, TextFontRegular } from "@fluentui/react-icons";
import { activity, describeError } from "../../lib/activity";
import { activeKit, addColor, moveColor, newKit, normalizeHex, type BrandKit, type BrandState, type ColorTarget } from "../../lib/brand";
import { applyColor, applyFonts, brandStore, colorFromSelection, fontFromSelection, kitFromDeck } from "../../lib/brandActions";
import { UserError } from "../../lib/ppt";
import { pane } from "../store";
import { Section, Seg, useUi } from "../ui";

const useStyles = makeStyles({
  swatches: { display: "grid", gridTemplateColumns: "repeat(4, minmax(0, 1fr))", gap: "8px" },
  swatchCell: { display: "flex", flexDirection: "column", alignItems: "center", gap: "2px", minWidth: 0 },
  swatch: {
    width: "100%",
    aspectRatio: "1.6",
    borderRadius: tokens.borderRadiusMedium,
    border: "1px solid rgba(0,0,0,.18)",
    cursor: "pointer",
    padding: 0,
    ":hover": { outline: `2px solid ${tokens.colorBrandStroke1}` },
    ":focus-visible": { outline: `2px solid ${tokens.colorStrokeFocus2}`, outlineOffset: "2px" },
  },
  label: { fontSize: tokens.fontSizeBase100, color: tokens.colorNeutralForeground3, fontFamily: "ui-monospace, Menlo, monospace" },
  editRow: { display: "flex", gap: "0" },
  grow: { flex: 1, minWidth: 0 },
});

const TARGETS: { value: ColorTarget; label: string }[] = [
  { value: "fill", label: "Fill" },
  { value: "line", label: "Outline" },
  { value: "text", label: "Text" },
];

async function run(name: string, work: () => Promise<string | void>) {
  pane.showMessage(undefined);
  try {
    const text = await activity.track(name, work);
    if (text) pane.showMessage({ intent: "success", text });
  } catch (e) {
    pane.showMessage({ intent: e instanceof UserError ? "warning" : "error", text: e instanceof UserError ? e.message : describeError(e) });
  }
}

export function Brand() {
  const s = useStyles();
  const ui = useUi();
  const [state, setState] = useState<BrandState | undefined>();
  const [target, setTarget] = useState<ColorTarget>("fill");
  const [editing, setEditing] = useState(false);
  const [hex, setHex] = useState("");
  const [naming, setNaming] = useState<"new" | "rename" | undefined>();
  const [name, setName] = useState("");
  const [confirmDelete, setConfirmDelete] = useState(false);

  useEffect(() => {
    brandStore
      .load()
      .then(setState)
      .catch(() => setState(undefined));
  }, []);

  if (!state) {
    return (
      <Section title="Brand kit">
        <Text size={200} className={ui.muted}>
          Loading…
        </Text>
      </Section>
    );
  }

  const kit = activeKit(state)!;
  const save = (next: BrandState) => {
    setState(next);
    void brandStore.save(next).catch((e) => pane.showMessage({ intent: "error", text: describeError(e) }));
  };
  const update = (patch: Partial<BrandKit>) => save({ ...state, kits: state.kits.map((k) => (k.id === kit.id ? { ...k, ...patch } : k)) });
  const addKit = (k: BrandKit) => save({ active: k.id, kits: [...state.kits, k] });

  const addHex = () => {
    const h = normalizeHex(hex);
    if (!h) {
      pane.showMessage({ intent: "warning", text: "Type a color as a hex code, like #C43E1C." });
      return;
    }
    update({ colors: addColor(kit.colors, h) });
    setHex("");
  };

  return (
    <>
      <Section title="Brand kit">
        {naming ? (
          <form
            className={ui.row}
            onSubmit={(e) => {
              e.preventDefault();
              if (naming === "new") addKit(newKit(name));
              else update({ name: name.trim() || kit.name });
              setNaming(undefined);
            }}
          >
            <Field label={naming === "new" ? "New kit name" : "Kit name"} className={s.grow}>
              <Input value={name} onChange={(_, d) => setName(d.value)} autoFocus />
            </Field>
            <Button appearance="primary" type="submit">
              {naming === "new" ? "Create" : "Save"}
            </Button>
            <Button onClick={() => setNaming(undefined)}>Cancel</Button>
          </form>
        ) : (
          <>
            <Field label="Kit">
              <Select value={kit.id} onChange={(_, d) => save({ ...state, active: d.value })}>
                {state.kits.map((k) => (
                  <option key={k.id} value={k.id}>
                    {k.name}
                  </option>
                ))}
              </Select>
            </Field>
            <div className={ui.row}>
              <Button
                size="small"
                icon={<AddRegular />}
                onClick={() => {
                  setName("");
                  setNaming("new");
                }}
              >
                New
              </Button>
              <Button
                size="small"
                icon={<EditRegular />}
                onClick={() => {
                  setName(kit.name);
                  setNaming("rename");
                }}
              >
                Rename
              </Button>
              {state.kits.length > 1 &&
                (confirmDelete ? (
                  <>
                    <Button
                      size="small"
                      appearance="primary"
                      onClick={() => {
                        const kits = state.kits.filter((k) => k.id !== kit.id);
                        save({ active: kits[0].id, kits });
                        setConfirmDelete(false);
                      }}
                    >
                      Delete “{kit.name}”
                    </Button>
                    <Button size="small" onClick={() => setConfirmDelete(false)}>
                      Keep
                    </Button>
                  </>
                ) : (
                  <Button size="small" icon={<DeleteRegular />} onClick={() => setConfirmDelete(true)}>
                    Delete
                  </Button>
                ))}
            </div>
            <Button
              size="small"
              appearance="subtle"
              icon={<DocumentArrowDownRegular />}
              style={{ alignSelf: "flex-start" }}
              onClick={() =>
                void run("pane brand import", async () => {
                  const k = await kitFromDeck();
                  addKit(k);
                  return `Made “${k.name}” with ${k.colors.length} colors${k.headingFont ? `, ${k.headingFont} for headings and ${k.bodyFont} for body text` : ""}.`;
                })
              }
            >
              Import colors and fonts from this deck
            </Button>
          </>
        )}
      </Section>

      <Section title="Colors">
        {kit.colors.length === 0 ? (
          <Text size={200} className={ui.muted}>
            No colors yet. Add them below, take one from a selected shape, or import this deck's theme above.
          </Text>
        ) : (
          <>
            <Seg label="Apply to" value={target} options={TARGETS} onChange={setTarget} />
            <div className={s.swatches}>
              {kit.colors.map((c, i) => (
                <div key={c} className={s.swatchCell}>
                  <button
                    className={s.swatch}
                    style={{ background: c }}
                    aria-label={`Apply ${c} to ${target === "line" ? "outline" : target}`}
                    title={i < 3 ? `${c} · ribbon Fill ${i + 1}` : c}
                    onClick={() => void run("pane brand color", async () => void (await applyColor(c, target)))}
                  />
                  {editing ? (
                    <div className={s.editRow}>
                      <Button size="small" appearance="transparent" icon={<ArrowLeftRegular />} aria-label={`Move ${c} left`} disabled={i === 0} onClick={() => update({ colors: moveColor(kit.colors, i, -1) })} />
                      <Button size="small" appearance="transparent" icon={<DismissRegular />} aria-label={`Remove ${c}`} onClick={() => update({ colors: kit.colors.filter((x) => x !== c) })} />
                      <Button
                        size="small"
                        appearance="transparent"
                        icon={<ArrowRightRegular />}
                        aria-label={`Move ${c} right`}
                        disabled={i === kit.colors.length - 1}
                        onClick={() => update({ colors: moveColor(kit.colors, i, 1) })}
                      />
                    </div>
                  ) : (
                    <span className={s.label}>{i < 3 ? `Fill ${i + 1}` : c}</span>
                  )}
                </div>
              ))}
            </div>
          </>
        )}
        {editing || kit.colors.length === 0 ? (
          <>
            <form
              className={ui.row}
              onSubmit={(e) => {
                e.preventDefault();
                addHex();
              }}
            >
              <Field label="Add a color" className={s.grow}>
                <Input value={hex} onChange={(_, d) => setHex(d.value)} placeholder="#C43E1C" />
              </Field>
              <Button type="submit">Add</Button>
            </form>
            <div className={ui.row}>
              <Button size="small" icon={<ColorRegular />} onClick={() => void run("pane brand pick color", async () => update({ colors: addColor(kit.colors, await colorFromSelection()) }))}>
                Add from selected shape
              </Button>
              {kit.colors.length > 0 && (
                <Button size="small" appearance="primary" onClick={() => setEditing(false)}>
                  Done
                </Button>
              )}
            </div>
          </>
        ) : (
          <Button size="small" appearance="subtle" icon={<EditRegular />} style={{ alignSelf: "flex-start" }} onClick={() => setEditing(true)}>
            Edit colors
          </Button>
        )}
        <Text size={200} className={ui.muted}>
          Click a color to apply it to the selected shapes. The first three are Fill 1, 2 and 3 on the ribbon.
        </Text>
      </Section>

      <Section title="Fonts">
        {(["headingFont", "bodyFont"] as const).map((key) => (
          <div key={key} className={ui.row}>
            <Field label={key === "headingFont" ? "Headings" : "Body text"} className={s.grow}>
              <Input value={kit[key]} onChange={(_, d) => update({ [key]: d.value })} placeholder={key === "headingFont" ? "e.g. Georgia" : "e.g. Arial"} />
            </Field>
            <Button
              icon={<TextFontRegular />}
              title="Use the selected text's font"
              aria-label={`Use the selected text's font for ${key === "headingFont" ? "headings" : "body text"}`}
              onClick={() => void run("pane brand pick font", async () => update({ [key]: await fontFromSelection() }))}
            />
          </div>
        ))}
        <div className={ui.grid2}>
          <Button onClick={() => void run("pane brand fonts", () => applyFonts("heading"))}>Heading font</Button>
          <Button onClick={() => void run("pane brand fonts", () => applyFonts("body"))}>Body font</Button>
        </div>
        <Button appearance="primary" onClick={() => void run("pane brand fonts", () => applyFonts())}>
          Apply brand fonts
        </Button>
        <Text size={200} className={ui.muted}>
          Heading font and Body font set the selected text. Apply brand fonts (also on the ribbon) gives titles the heading font and everything else the body font, on the selection or, with nothing
          selected, the whole slide.
        </Text>
      </Section>
    </>
  );
}
