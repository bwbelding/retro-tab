import { useCallback, useEffect, useMemo, useState } from "react";
import {
  Badge,
  Button,
  Field,
  Input,
  Menu,
  MenuItem,
  MenuList,
  MenuPopover,
  MenuTrigger,
  Text,
  ToggleButton,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { AddRegular, MoreHorizontalRegular, SearchRegular, ShapesRegular, StarFilled, StarRegular } from "@fluentui/react-icons";
import { activity, describeError } from "../../lib/activity";
import { categoriesOf, filterItems, type Filter, type LibraryItem } from "../../lib/library";
import { cleanupHelperSlides, insertItem, library, saveSelection, stripOldItems } from "../../lib/libraryActions";
import { UserError } from "../../lib/ppt";
import { pane } from "../store";
import { Section, useUi } from "../ui";
import { ShapeTest } from "./ShapeTest";

const useStyles = makeStyles({
  grid: { display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: "8px" },
  card: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    padding: "6px",
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    minWidth: 0,
  },
  thumb: {
    height: "72px",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    borderRadius: tokens.borderRadiusSmall,
    // A mid-gray checkerboard, so both dark shapes and white ones (for dark slides) show up.
    backgroundColor: "#b4b4b4",
    backgroundImage: "linear-gradient(45deg, #a2a2a2 25%, transparent 25%, transparent 75%, #a2a2a2 75%), linear-gradient(45deg, #a2a2a2 25%, transparent 25%, transparent 75%, #a2a2a2 75%)",
    backgroundSize: "12px 12px",
    backgroundPosition: "0 0, 6px 6px",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    padding: "4px",
    cursor: "pointer",
    ":hover": { filter: "brightness(1.06)" },
    ":focus-visible": { outline: `2px solid ${tokens.colorStrokeFocus2}` },
  },
  img: { maxWidth: "100%", maxHeight: "100%", objectFit: "contain" },
  name: { fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  meta: { display: "flex", alignItems: "center", gap: "2px", minWidth: 0 },
  grow: { flex: 1, minWidth: 0 },
  chips: { display: "flex", gap: "4px", flexWrap: "wrap" },
  form: { display: "flex", flexDirection: "column", gap: "8px" },
});

function report(e: unknown) {
  pane.showMessage({ intent: e instanceof UserError ? "warning" : "error", text: e instanceof UserError ? e.message : describeError(e) });
}

/** Name and category fields, for saving and for editing an item. */
function ItemForm({
  initial,
  categories,
  submit,
  cancel,
  label,
}: {
  initial: { name: string; category: string };
  categories: string[];
  submit: (name: string, category: string) => Promise<void>;
  cancel: () => void;
  label: string;
}) {
  const s = useStyles();
  const ui = useUi();
  const [name, setName] = useState(initial.name);
  const [category, setCategory] = useState(initial.category);
  const [busy, setBusy] = useState(false);
  const go = async () => {
    setBusy(true);
    try {
      await submit(name, category);
    } finally {
      setBusy(false);
    }
  };
  return (
    <form
      className={s.form}
      onSubmit={(e) => {
        e.preventDefault();
        void go();
      }}
    >
      <Field label="Name">
        <Input value={name} onChange={(_, d) => setName(d.value)} autoFocus />
      </Field>
      <Field label="Category">
        <Input value={category} onChange={(_, d) => setCategory(d.value)} list="retro-categories" placeholder="e.g. Callouts" />
      </Field>
      <datalist id="retro-categories">
        {categories.map((c) => (
          <option key={c} value={c} />
        ))}
      </datalist>
      <div className={ui.row}>
        <Button appearance="primary" type="submit" disabled={busy}>
          {label}
        </Button>
        <Button onClick={cancel} disabled={busy}>
          Cancel
        </Button>
      </div>
    </form>
  );
}

function ItemCard({ item, categories, reload }: { item: LibraryItem; categories: string[]; reload: () => void }) {
  const s = useStyles();
  const ui = useUi();
  const [mode, setMode] = useState<"view" | "edit" | "delete">("view");

  const insert = async () => {
    pane.showMessage(undefined);
    try {
      const route = await activity.track("pane insert shape", () => insertItem(item));
      if (route === "helper") {
        pane.showMessage({
          intent: "info",
          text: `“${item.name}” is on a helper slide, selected. Press ⌘X; Retro removes the helper slide and takes you back to your slide, where you press ⌘V.`,
        });
      }
      reload();
    } catch (e) {
      report(e);
    }
  };

  if (mode === "edit") {
    return (
      <div className={s.card} style={{ gridColumn: "1 / -1" }}>
        <ItemForm
          label="Save"
          initial={item}
          categories={categories}
          cancel={() => setMode("view")}
          submit={async (name, category) => {
            await library.update(item.id, { name: name.trim() || item.name, category: category.trim() });
            setMode("view");
            reload();
          }}
        />
      </div>
    );
  }

  return (
    <div className={s.card}>
      <button className={s.thumb} onClick={() => void insert()} title={`Insert ${item.name}`} aria-label={`Insert ${item.name}`}>
        {item.preview ? <img className={s.img} src={item.preview} alt="" /> : <ShapesRegular fontSize={32} />}
      </button>
      {mode === "delete" ? (
        <div className={s.form}>
          <Text size={200}>Delete “{item.name}”?</Text>
          <div className={ui.row}>
            <Button
              size="small"
              appearance="primary"
              onClick={() =>
                void library
                  .remove(item.id)
                  .then(reload)
                  .catch(report)
              }
            >
              Delete
            </Button>
            <Button size="small" onClick={() => setMode("view")}>
              Keep
            </Button>
          </div>
        </div>
      ) : (
        <div className={s.meta}>
          <div className={s.grow}>
            <div className={s.name} title={item.name}>
              {item.name}
            </div>
            {item.route === "helper" ? (
              <Badge size="small" appearance="tint" color="warning">
                Helper slide
              </Badge>
            ) : (
              <Badge size="small" appearance="tint" color="success">
                One click
              </Badge>
            )}
          </div>
          <Button
            size="small"
            appearance="outline"
            icon={item.favorite ? <StarFilled /> : <StarRegular />}
            aria-label={item.favorite ? "Remove from favorites" : "Add to favorites"}
            aria-pressed={item.favorite}
            onClick={() =>
              void library
                .update(item.id, { favorite: !item.favorite })
                .then(reload)
                .catch(report)
            }
          />
          <Menu>
            <MenuTrigger disableButtonEnhancement>
              <Button size="small" appearance="outline" icon={<MoreHorizontalRegular />} aria-label={`More for ${item.name}`} />
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                <MenuItem onClick={() => void insert()}>Insert</MenuItem>
                {item.route === "helper" && (
                  <MenuItem onClick={() => pane.showMessage({ intent: "info", text: `“${item.name}” uses a helper slide because PowerPoint won't let Retro recreate: ${item.issues.join(", ")}.` })}>
                    Why a helper slide?
                  </MenuItem>
                )}
                <MenuItem onClick={() => setMode("edit")}>Rename or move…</MenuItem>
                <MenuItem onClick={() => setMode("delete")}>Delete…</MenuItem>
              </MenuList>
            </MenuPopover>
          </Menu>
        </div>
      )}
    </div>
  );
}

export function Shapes() {
  const s = useStyles();
  const ui = useUi();
  const [items, setItems] = useState<LibraryItem[] | undefined>();
  const [saving, setSaving] = useState(false);
  const [query, setQuery] = useState("");
  const [filter, setFilter] = useState<Filter>("all");
  const [checking, setChecking] = useState(false);

  const reload = useCallback(() => {
    library
      .list()
      .then(setItems)
      .catch(() => setItems([]));
  }, []);

  useEffect(() => {
    // Load the library, tidy leftover helper slides, and keep tidying as the selection changes
    // (cutting the shapes off a helper slide is a selection change).
    reload();
    void stripOldItems().catch(() => undefined);
    if (typeof PowerPoint === "undefined") return;
    const tidy = () => void cleanupHelperSlides().catch(() => undefined);
    tidy();
    const doc = Office.context?.document;
    if (!doc?.addHandlerAsync) return;
    doc.addHandlerAsync(Office.EventType.DocumentSelectionChanged, tidy);
    return () => doc.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: tidy });
  }, [reload]);

  const categories = useMemo(() => categoriesOf(items ?? []), [items]);
  const shown = useMemo(() => filterItems(items ?? [], query, filter), [items, query, filter]);
  const chips: { value: Filter; label: string }[] = [
    { value: "all", label: "All" },
    { value: "favorites", label: "★ Favorites" },
    ...categories.map((c) => ({ value: `category:${c}` as Filter, label: c })),
  ];

  return (
    <>
      <Section title="Shape library">
        {saving ? (
          <ItemForm
            label="Save to library"
            initial={{ name: "", category: filter.startsWith("category:") ? filter.slice("category:".length) : "" }}
            categories={categories}
            cancel={() => setSaving(false)}
            submit={async (name, category) => {
              try {
                const item = await activity.track("pane save shape", () => saveSelection(name, category));
                setSaving(false);
                pane.showMessage({
                  intent: "success",
                  text:
                    item.route === "click"
                      ? `Saved “${item.name}”. It inserts in one click.`
                      : `Saved “${item.name}”. It inserts with a helper slide, because of: ${item.issues.join(", ")}.`,
                });
                reload();
              } catch (e) {
                report(e);
              }
            }}
          />
        ) : (
          <Button appearance="primary" icon={<AddRegular />} style={{ alignSelf: "flex-start" }} onClick={() => setSaving(true)}>
            Save selection to library
          </Button>
        )}
      </Section>

      <Section>
        {items === undefined ? (
          <Text size={200} className={ui.muted}>
            Loading…
          </Text>
        ) : items.length === 0 ? (
          <Text size={200} className={ui.muted}>
            Your library is empty. Select shapes on a slide, then click “Save selection to library”. Click a saved shape to insert it where it was saved.
          </Text>
        ) : (
          <>
            <Input contentBefore={<SearchRegular />} placeholder="Search" value={query} onChange={(_, d) => setQuery(d.value)} aria-label="Search the library" />
            <div className={s.chips} role="radiogroup" aria-label="Show">
              {chips.map((c) => (
                <ToggleButton key={c.value} size="small" role="radio" aria-checked={c.value === filter} checked={c.value === filter} onClick={() => setFilter(c.value)}>
                  {c.label}
                </ToggleButton>
              ))}
            </div>
            {shown.length === 0 ? (
              <Text size={200} className={ui.muted}>
                Nothing matches.
              </Text>
            ) : (
              <div className={s.grid}>
                {shown.map((item) => (
                  <ItemCard key={item.id} item={item} categories={categories} reload={reload} />
                ))}
              </div>
            )}
            <Text size={200} className={ui.muted}>
              Most recently used first. Click a shape to insert it where it was saved. <b>One click</b> shapes appear on your slide. <b>Helper slide</b> shapes appear on a new slide, selected:
              press ⌘X and Retro takes you back to your slide to press ⌘V.
            </Text>
          </>
        )}
      </Section>

      {checking ? (
        <ShapeTest />
      ) : (
        <Section>
          <Button size="small" style={{ alignSelf: "flex-start" }} onClick={() => setChecking(true)}>
            Check what Retro can copy from a shape…
          </Button>
        </Section>
      )}
    </>
  );
}
