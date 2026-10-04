import { useCallback, useEffect, useMemo, useState } from "react";
import { Button, Checkbox, Input, Spinner, Text, ToggleButton, makeStyles, tokens } from "@fluentui/react-components";
import { ArrowClockwiseRegular, CheckmarkCircleRegular, SettingsRegular, WrenchRegular } from "@fluentui/react-icons";
import { activity, describeError } from "../../lib/activity";
import { activeKit } from "../../lib/brand";
import { brandStore } from "../../lib/brandActions";
import { autoFixable, checkDeck, countByRule, RULES, type Issue, type Kit, type Rule } from "../../lib/deckCheck";
import { applyFixes, goToShape, snapshotDeck } from "../../lib/deckCheckActions";
import { prefs } from "../../lib/prefs";
import { pane, usePane } from "../store";
import { Section, Seg, useUi } from "../ui";
import { FindUpdate } from "./FindUpdate";
import { Storyline } from "./Storyline";

const useStyles = makeStyles({
  header: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: "8px" },
  chips: { display: "flex", gap: "4px", flexWrap: "wrap" },
  slide: { fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold, color: tokens.colorNeutralForeground2, marginTop: "4px" },
  issue: {
    display: "flex",
    gap: "8px",
    alignItems: "flex-start",
    padding: "8px",
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  body: { flex: 1, minWidth: 0, display: "flex", flexDirection: "column", gap: "2px" },
  line: { fontSize: tokens.fontSizeBase200, lineHeight: tokens.lineHeightBase200 },
  value: { fontFamily: "ui-monospace, Menlo, monospace", fontSize: tokens.fontSizeBase200, padding: "0 4px", borderRadius: tokens.borderRadiusSmall, backgroundColor: tokens.colorNeutralBackground3 },
  shape: { fontSize: tokens.fontSizeBase100, color: tokens.colorNeutralForeground3, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  swatch: { display: "inline-block", width: "10px", height: "10px", borderRadius: "2px", border: "1px solid rgba(0,0,0,.2)", verticalAlign: "-1px", marginRight: "4px" },
  ok: { display: "flex", alignItems: "center", gap: "6px", color: tokens.colorPaletteGreenForeground1 },
  rules: { display: "flex", flexDirection: "column", gap: "2px" },
  buttons: { display: "flex", flexDirection: "column", gap: "4px", alignItems: "stretch" },
  alt: { display: "flex", gap: "4px", marginTop: "4px" },
});

type Filter = "all" | Rule;

interface Result {
  issues: Issue[];
  slides: number;
  full: boolean;
  kit?: Kit;
}

/** The Check pane: the deck check, or find and update. */
export function Check() {
  const { checkMode } = usePane();
  return (
    <>
      <Section>
        <Seg
          label="Check, find or storyline"
          value={checkMode}
          options={[
            { value: "check", label: "Deck check" },
            { value: "find", label: "Find" },
            { value: "story", label: "Storyline" },
          ]}
          onChange={pane.setCheckMode}
        />
      </Section>
      {checkMode === "find" ? <FindUpdate /> : checkMode === "story" ? <Storyline /> : <DeckCheck />}
    </>
  );
}

function DeckCheck() {
  const s = useStyles();
  const ui = useUi();
  const [rules, setRules] = useState<Rule[] | undefined>();
  const [result, setResult] = useState<Result | undefined>();
  const [running, setRunning] = useState(false);
  const [filter, setFilter] = useState<Filter>("all");
  const [showRules, setShowRules] = useState(false);

  const run = useCallback(async (on: Rule[]) => {
    if (typeof PowerPoint === "undefined") return;
    setRunning(true);
    pane.showMessage(undefined);
    try {
      const [snapshot, state] = await Promise.all([activity.track("pane deck check", snapshotDeck), brandStore.load()]);
      const kit = activeKit(state);
      setResult({ issues: checkDeck(snapshot.slides, kit, snapshot.size, new Set(on)), slides: snapshot.slides.length, full: snapshot.full, kit });
    } catch (e) {
      pane.showMessage({ intent: "error", text: `Deck check didn't finish: ${describeError(e)}` });
    } finally {
      setRunning(false);
    }
  }, []);

  useEffect(() => {
    // Check as soon as the pane opens.
    void prefs.checkRules().then((on) => {
      setRules(on);
      void run(on);
    });
  }, [run]);

  const counts = useMemo(() => countByRule(result?.issues ?? []), [result]);
  const shown = useMemo(() => (result?.issues ?? []).filter((i) => filter === "all" || i.rule === filter), [result, filter]);
  const bySlide = useMemo(() => {
    const groups: { key: string; title: string; issues: Issue[] }[] = [];
    for (const issue of shown) {
      const last = groups[groups.length - 1];
      if (last?.key === issue.slideId) last.issues.push(issue);
      else groups.push({ key: issue.slideId, title: `Slide ${issue.slideIndex + 1}${issue.slideTitle ? ` · ${issue.slideTitle}` : ""}`, issues: [issue] });
    }
    return groups;
  }, [shown]);

  const toggleRule = (rule: Rule, on: boolean) => {
    const next = RULES.map((r) => r.rule).filter((r) => (r === rule ? on : rules?.includes(r)));
    setRules(next);
    void prefs.setCheckRules(next);
    if (filter === rule && !on) setFilter("all");
  };

  const plural = (n: number, word: string) => `${n} ${word}${n === 1 ? "" : "s"}`;
  const kit = result?.kit;
  const fixable = autoFixable(shown);

  /** Fix one issue and drop it from the list (the rest of the results still hold). */
  const fixOne = async (issue: Issue, alt?: string) => {
    try {
      await activity.track("pane deck fix", () => applyFixes([{ slideId: issue.slideId, fix: issue.fix!, alt }]));
      setResult((r) => r && { ...r, issues: r.issues.filter((i) => i !== issue) });
    } catch (e) {
      pane.showMessage({ intent: "warning", text: `Couldn't fix that. The slide may have changed: run the check again. (${describeError(e)})` });
    }
  };

  /** Fix every issue shown that doesn't need your words, then check again. */
  const fixAll = async () => {
    let message: Parameters<typeof pane.showMessage>[0];
    try {
      const n = await activity.track("pane deck fix all", () => applyFixes(fixable.map((i) => ({ slideId: i.slideId, fix: i.fix! }))));
      message = { intent: "success", text: `Fixed ${plural(n, "issue")}. ⌘Z undoes them.` };
    } catch (e) {
      message = { intent: "warning", text: `Couldn't fix everything. The slides may have changed: run the check again. (${describeError(e)})` };
    }
    // Checking again clears the pane's message, so it's shown after.
    if (rules) await run(rules);
    pane.showMessage(message);
  };

  return (
    <>
      <Section>
        <div className={s.header}>
          <div>
            <Text weight="semibold">{running ? "Checking…" : result ? (result.issues.length ? plural(result.issues.length, "issue") : "Ready to send") : "Deck check"}</Text>
            {result && !running && (
              <Text size={200} className={ui.muted}>
                {" "}
                · {plural(result.slides, "slide")} checked
              </Text>
            )}
          </div>
          <Button size="small" style={{ flexShrink: 0 }} icon={running ? <Spinner size="extra-tiny" /> : <ArrowClockwiseRegular />} disabled={running || !rules} onClick={() => rules && void run(rules)}>
            Run again
          </Button>
        </div>

        {!running && fixable.length > 0 && (
          <Button appearance="primary" icon={<WrenchRegular />} style={{ alignSelf: "flex-start" }} onClick={() => void fixAll()}>
            {fixable.length === 1 ? "Fix 1 issue" : `Fix all ${fixable.length}${filter === "all" ? "" : ` ${RULES.find((r) => r.rule === filter)!.label.toLowerCase()}`}`}
          </Button>
        )}

        {result && result.issues.length > 0 && (
          <div className={s.chips} role="radiogroup" aria-label="Show">
            <ToggleButton size="small" role="radio" aria-checked={filter === "all"} checked={filter === "all"} onClick={() => setFilter("all")}>
              All {result.issues.length}
            </ToggleButton>
            {RULES.filter((r) => counts[r.rule] > 0).map((r) => (
              <ToggleButton key={r.rule} size="small" role="radio" aria-checked={filter === r.rule} checked={filter === r.rule} onClick={() => setFilter(r.rule)}>
                {r.label} {counts[r.rule]}
              </ToggleButton>
            ))}
          </div>
        )}

        {result && !running && result.issues.length === 0 && (
          <div className={s.ok}>
            <CheckmarkCircleRegular fontSize={20} />
            <Text>No issues on {plural(result.slides, "slide")}. Good to send.</Text>
          </div>
        )}

        {bySlide.map((group) => (
          <div key={group.key} className={s.rules}>
            <div className={s.slide}>{group.title}</div>
            {group.issues.map((issue, n) => (
              <div key={n} className={s.issue}>
                <div className={s.body}>
                  <div className={s.line}>
                    <b>{issue.label}</b> {issue.value && <span className={s.value}>{issue.value}</span>} {issue.text}
                  </div>
                  {issue.closest && (
                    <div className={s.line}>
                      Closest: <span className={s.swatch} style={{ background: issue.closest }} />
                      <span className={s.value}>{issue.closest}</span>
                    </div>
                  )}
                  <div className={s.shape} title={issue.shape}>
                    {issue.shape}
                  </div>
                  {issue.fix?.kind === "alt" && <AltText onSave={(text) => void fixOne(issue, text)} />}
                </div>
                <div className={s.buttons}>
                  <Button
                    size="small"
                    onClick={() =>
                      void goToShape(issue.slideId, issue.shapeId).catch((e) => pane.showMessage({ intent: "warning", text: `Couldn't go there. The shape may have changed: run the check again. (${describeError(e)})` }))
                    }
                  >
                    Go to
                  </Button>
                  {issue.fix && issue.fixLabel && (
                    <Button size="small" appearance="outline" onClick={() => void fixOne(issue)}>
                      {issue.fixLabel}
                    </Button>
                  )}
                </div>
              </div>
            ))}
          </div>
        ))}

        {result && !result.full && rules?.includes("alttext") && (
          <Text size={200} className={ui.muted}>
            The alt text and empty placeholder checks need PowerPoint 16.105 or later, and this Mac has {Office.context.diagnostics?.version ?? "an older version"}.
          </Text>
        )}
      </Section>

      <Section>
        <div className={s.header}>
          <Text size={200} className={ui.muted}>
            {kit ? `Checks against kit “${kit.name}”` : "No brand kit yet: fonts and colors aren't checked."}
          </Text>
          <Button size="small" icon={<SettingsRegular />} aria-expanded={showRules} onClick={() => setShowRules(!showRules)}>
            Rules
          </Button>
        </div>
        {kit && (!kit.colors.length || !(kit.headingFont || kit.bodyFont)) && (
          <Text size={200} className={ui.muted}>
            {!kit.colors.length && "Add colors to the kit to check colors. "}
            {!(kit.headingFont || kit.bodyFont) && "Add fonts to the kit to check fonts."}
          </Text>
        )}
        {showRules && rules && (
          <div className={s.rules}>
            {RULES.map((r) => (
              <Checkbox key={r.rule} label={`${r.label}: ${r.help}`} checked={rules.includes(r.rule)} onChange={(_, d) => toggleRule(r.rule, Boolean(d.checked))} />
            ))}
            <Text size={200} className={ui.muted}>
              Changes apply the next time you run the check. Retro's own notes, stamps and tracker are never flagged.
            </Text>
          </div>
        )}
      </Section>
    </>
  );
}

/** Type a picture's alt text right in the result. */
function AltText({ onSave }: { onSave: (text: string) => void }) {
  const s = useStyles();
  const [text, setText] = useState("");
  return (
    <form
      className={s.alt}
      onSubmit={(e) => {
        e.preventDefault();
        if (text.trim()) onSave(text);
      }}
    >
      <Input size="small" value={text} onChange={(_, d) => setText(d.value)} placeholder="Describe the picture" aria-label="Alt text" style={{ flex: 1, minWidth: 0 }} />
      <Button size="small" appearance="outline" type="submit" disabled={!text.trim()}>
        Save
      </Button>
    </form>
  );
}
