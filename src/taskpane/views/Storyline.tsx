import { useCallback, useEffect, useState } from "react";
import { Badge, Button, Input, Text, makeStyles, tokens } from "@fluentui/react-components";
import { ArrowClockwiseRegular } from "@fluentui/react-icons";
import { activity, describeError } from "../../lib/activity";
import { goToShape } from "../../lib/deckCheckActions";
import { MIN_TITLE_WORDS, readTitles, setTitle, titleState, type TitleRow } from "../../lib/storyline";
import { pane } from "../store";
import { Section, useUi } from "../ui";

const useStyles = makeStyles({
  header: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: "8px" },
  row: { display: "grid", gridTemplateColumns: "24px minmax(0, 1fr) auto", gap: "6px", alignItems: "start", padding: "6px 0", borderBottom: `1px solid ${tokens.colorNeutralStroke3}` },
  number: { fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3, paddingTop: "6px", textAlign: "right" },
  body: { display: "flex", flexDirection: "column", gap: "4px", minWidth: 0 },
  missing: { fontSize: tokens.fontSizeBase200, color: tokens.colorNeutralForeground3, fontStyle: "italic", paddingTop: "6px" },
});

/** One title, edited in place: saved when you press Return or leave the field. */
function TitleField({ row, saved }: { row: TitleRow; saved: () => void }) {
  const [text, setText] = useState(row.title);
  const save = async () => {
    const next = text.replace(/\s+/g, " ").trim();
    if (!row.shapeId || next === row.title) return;
    try {
      await activity.track("pane storyline edit", () => setTitle(row.slideId, row.shapeId!, next));
      saved();
    } catch (e) {
      pane.showMessage({ intent: "warning", text: `Couldn't change that title: ${describeError(e)}` });
    }
  };
  return (
    <Input
      size="small"
      value={text}
      aria-label={`Title of slide ${row.index + 1}`}
      onChange={(_, d) => setText(d.value)}
      onBlur={() => void save()}
      onKeyDown={(e) => {
        if (e.key === "Enter") (e.target as HTMLInputElement).blur();
      }}
    />
  );
}

export function Storyline() {
  const s = useStyles();
  const ui = useUi();
  const [rows, setRows] = useState<TitleRow[] | undefined>();
  const [loading, setLoading] = useState(false);

  const load = useCallback(async () => {
    if (typeof PowerPoint === "undefined") return;
    setLoading(true);
    try {
      setRows(await activity.track("pane storyline", readTitles));
    } catch (e) {
      pane.showMessage({ intent: "error", text: `Couldn't read the titles: ${describeError(e)}` });
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    // Read the titles as soon as the view opens.
    if (typeof PowerPoint === "undefined") return;
    readTitles()
      .then(setRows)
      .catch((e) => pane.showMessage({ intent: "error", text: `Couldn't read the titles: ${describeError(e)}` }));
  }, []);

  const short = rows?.filter((r) => titleState(r) === "short").length ?? 0;
  const missing = rows?.filter((r) => titleState(r) === "missing").length ?? 0;

  return (
    <Section title="Storyline">
      <div className={s.header}>
        <Text size={200} className={ui.muted}>
          {rows ? `${rows.length} slides · ${short} topic label${short === 1 ? "" : "s"} · ${missing} without a title` : "Reading titles…"}
        </Text>
        <Button size="small" style={{ flexShrink: 0 }} icon={<ArrowClockwiseRegular />} disabled={loading} onClick={() => void load()}>
          Refresh
        </Button>
      </div>
      <Text size={200} className={ui.muted}>
        Read the titles alone: they should tell the whole story. Titles under {MIN_TITLE_WORDS} words usually name a topic (“Market size”) instead of making the point (“Market grew 30% in Q3”).
      </Text>
      {rows?.map((row) => {
        const state = titleState(row);
        return (
          <div key={row.slideId} className={s.row}>
            <span className={s.number}>{row.index + 1}</span>
            <div className={s.body}>
              {row.shapeId ? (
                // Keyed by title so a refresh shows the deck's current text.
                <TitleField key={row.title} row={row} saved={() => void load()} />
              ) : (
                <span className={s.missing}>No title placeholder on this slide</span>
              )}
              {state === "short" && (
                <Badge size="small" appearance="tint" color="warning">
                  Topic label
                </Badge>
              )}
              {state === "missing" && row.shapeId && (
                <Badge size="small" appearance="tint" color="danger">
                  No title
                </Badge>
              )}
            </div>
            <Button size="small" onClick={() => void goToShape(row.slideId, "").catch(() => undefined)}>
              Go to
            </Button>
          </div>
        );
      })}
    </Section>
  );
}
