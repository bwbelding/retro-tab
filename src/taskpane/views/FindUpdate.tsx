import { useMemo, useState } from "react";
import { Button, Checkbox, Field, Input, Text, makeStyles, tokens } from "@fluentui/react-components";
import { SearchRegular } from "@fluentui/react-icons";
import { activity, describeError } from "../../lib/activity";
import { goToShape } from "../../lib/deckCheckActions";
import { findMatches, type Match, type TextSource } from "../../lib/findReplace";
import { collectText, replaceMatches } from "../../lib/findActions";
import { pane } from "../store";
import { Section, useUi } from "../ui";

const useStyles = makeStyles({
  form: { display: "flex", flexDirection: "column", gap: "8px" },
  options: { display: "flex", gap: "4px", flexWrap: "wrap" },
  slide: { fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold, color: tokens.colorNeutralForeground2, marginTop: "4px" },
  match: { display: "flex", gap: "4px", alignItems: "flex-start", padding: "4px 4px 4px 0", borderBottom: `1px solid ${tokens.colorNeutralStroke3}` },
  body: { flex: 1, minWidth: 0, display: "flex", flexDirection: "column", gap: "2px", paddingTop: "6px" },
  context: { fontSize: tokens.fontSizeBase200, lineHeight: tokens.lineHeightBase200, wordBreak: "break-word" },
  found: { backgroundColor: tokens.colorPaletteYellowBackground2, color: tokens.colorNeutralForeground1, fontWeight: tokens.fontWeightSemibold, borderRadius: "2px", padding: "0 1px" },
  where: { fontSize: tokens.fontSizeBase100, color: tokens.colorNeutralForeground3, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  footer: { display: "flex", gap: "8px", alignItems: "center", justifyContent: "space-between" },
});

export function FindUpdate() {
  const s = useStyles();
  const ui = useUi();
  const [query, setQuery] = useState("");
  const [replacement, setReplacement] = useState("");
  const [matchCase, setMatchCase] = useState(false);
  const [wholeWord, setWholeWord] = useState(true);
  const [sources, setSources] = useState<TextSource[] | undefined>();
  const [searched, setSearched] = useState("");
  const [skip, setSkip] = useState<Set<string>>(new Set());
  const [busy, setBusy] = useState(false);

  const matches = useMemo(() => (sources ? findMatches(sources, searched, { matchCase, wholeWord }) : []), [sources, searched, matchCase, wholeWord]);
  const chosen = matches.filter((m) => !skip.has(m.key));
  const bySlide = useMemo(() => {
    const groups: { key: string; index: number; matches: Match[] }[] = [];
    for (const m of matches) {
      const last = groups[groups.length - 1];
      if (last?.key === m.source.slideId) last.matches.push(m);
      else groups.push({ key: m.source.slideId, index: m.source.slideIndex, matches: [m] });
    }
    return groups;
  }, [matches]);

  const plural = (n: number, word: string) => `${n} ${word}${n === 1 ? "" : "es"}`;

  const search = async () => {
    if (!query) return;
    setBusy(true);
    pane.showMessage(undefined);
    try {
      setSources(await activity.track("pane find", collectText));
      setSearched(query);
      setSkip(new Set());
    } catch (e) {
      pane.showMessage({ intent: "error", text: `Couldn't read the deck: ${describeError(e)}` });
    } finally {
      setBusy(false);
    }
  };

  const replace = async () => {
    setBusy(true);
    try {
      const { replaced, skipped } = await activity.track("pane replace", () => replaceMatches(chosen, replacement));
      // Read the deck again so the list shows what's left.
      setSources(await collectText());
      setSkip(new Set());
      pane.showMessage({
        intent: skipped ? "warning" : "success",
        text: `Replaced ${plural(replaced, "match")}${replacement ? ` with “${replacement}”` : ""}.${skipped ? ` ${plural(skipped, "match")} left alone because their text changed since you searched: search again to see them.` : ""} ⌘Z undoes this.`,
      });
    } catch (e) {
      pane.showMessage({ intent: "error", text: `Couldn't replace: ${describeError(e)}` });
    } finally {
      setBusy(false);
    }
  };

  const slides = new Set(matches.map((m) => m.source.slideId)).size;

  return (
    <Section title="Find & update">
      <form
        className={s.form}
        onSubmit={(e) => {
          e.preventDefault();
          void search();
        }}
      >
        <Field label="Find">
          <Input value={query} onChange={(_, d) => setQuery(d.value)} placeholder="e.g. $4.2M or Q3 FY26" contentBefore={<SearchRegular />} />
        </Field>
        <Field label="Replace with">
          <Input value={replacement} onChange={(_, d) => setReplacement(d.value)} placeholder="e.g. $4.5M" />
        </Field>
        <div className={s.options}>
          <Checkbox label="Match case" checked={matchCase} onChange={(_, d) => setMatchCase(Boolean(d.checked))} />
          <Checkbox label="Whole words and numbers" checked={wholeWord} onChange={(_, d) => setWholeWord(Boolean(d.checked))} />
        </div>
        <Button type="submit" disabled={!query || busy} style={{ alignSelf: "flex-start" }}>
          Find in deck
        </Button>
      </form>

      {sources && searched && (
        <>
          <Text size={200} weight="semibold">
            {matches.length ? `${plural(matches.length, "match")} on ${slides} slide${slides === 1 ? "" : "s"} for “${searched}”` : `No matches for “${searched}”.`}
          </Text>
          {bySlide.map((group) => (
            <div key={group.key}>
              <div className={s.slide}>Slide {group.index + 1}</div>
              {group.matches.map((m) => (
                <div key={m.key} className={s.match}>
                  <Checkbox
                    checked={!skip.has(m.key)}
                    aria-label={`Replace ${m.found} on slide ${m.source.slideIndex + 1}`}
                    onChange={(_, d) => {
                      const next = new Set(skip);
                      if (d.checked) next.delete(m.key);
                      else next.add(m.key);
                      setSkip(next);
                    }}
                  />
                  <div className={s.body}>
                    <div className={s.context}>
                      {m.before}
                      <span className={s.found}>{m.found}</span>
                      {m.after}
                    </div>
                    <div className={s.where} title={m.source.where}>
                      {m.source.where}
                    </div>
                  </div>
                  <Button size="small" style={{ marginTop: "4px" }} onClick={() => void goToShape(m.source.slideId, m.source.path[0]).catch(() => undefined)}>
                    Go to
                  </Button>
                </div>
              ))}
            </div>
          ))}
          {matches.length > 0 && (
            <div className={s.footer}>
              <Text size={200} className={ui.muted}>
                {chosen.length} of {matches.length} ticked
              </Text>
              <Button appearance="primary" disabled={busy || chosen.length === 0} onClick={() => void replace()}>
                {replacement ? `Replace ${chosen.length}` : `Remove ${chosen.length}`}
              </Button>
            </div>
          )}
          <Text size={200} className={ui.muted}>
            Searches every slide's text, including grouped shapes and tables. Replacing keeps the text's formatting, except inside table cells, which are rewritten whole: mixed formatting within a cell may be lost.
          </Text>
        </>
      )}
    </Section>
  );
}
