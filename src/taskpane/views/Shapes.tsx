import { useState } from "react";
import { Badge, Button, Text, makeStyles, tokens } from "@fluentui/react-components";
import { CopyRegular, PlayRegular, SlideAddRegular } from "@fluentui/react-icons";
import { activity, describeError } from "../../lib/activity";
import { UserError } from "../../lib/ppt";
import { formatShapeTest, runHelperSlideTest, runShapeTest, type ShapeTestResult } from "../../lib/shapeTest";
import { pane } from "../store";
import { Section, useUi } from "../ui";
import { StatusIcon } from "./Diagnostics";

const useStyles = makeStyles({
  row: {
    display: "grid",
    gridTemplateColumns: "18px minmax(0, 1fr)",
    gap: "8px",
    padding: "6px 0",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    fontSize: tokens.fontSizeBase200,
    ":first-child": { borderTop: "none" },
  },
  item: { display: "flex", flexDirection: "column", gap: "2px", padding: "6px 0", borderTop: `1px solid ${tokens.colorNeutralStroke2}`, fontSize: tokens.fontSizeBase200 },
  preview: { maxWidth: "100%", maxHeight: "120px", alignSelf: "flex-start", border: `1px solid ${tokens.colorNeutralStroke2}`, background: "#fff" },
  list: { margin: 0, paddingLeft: "18px", fontSize: tokens.fontSizeBase200 },
  report: { width: "100%", height: "120px", fontFamily: "ui-monospace, Menlo, monospace", fontSize: "11px" },
});

function fail(e: unknown) {
  pane.showMessage({ intent: e instanceof UserError ? "warning" : "error", text: e instanceof UserError ? e.message : describeError(e) });
}

export function Shapes() {
  const s = useStyles();
  const ui = useUi();
  const [result, setResult] = useState<ShapeTestResult | undefined>();
  const [busy, setBusy] = useState(false);
  const [copied, setCopied] = useState<"" | "copied" | "manual">("");

  const test = async () => {
    setBusy(true);
    setCopied("");
    setResult({ steps: [], items: [], differences: [] });
    pane.showMessage(undefined);
    try {
      await activity.track("pane shape test", () => runShapeTest(setResult));
    } catch (e) {
      fail(e);
    } finally {
      setBusy(false);
    }
  };

  const helper = async () => {
    setBusy(true);
    try {
      const text = await activity.track("pane helper slide test", runHelperSlideTest);
      pane.showMessage({ intent: "success", text });
    } catch (e) {
      fail(e);
    } finally {
      setBusy(false);
    }
  };

  const report = result ? formatShapeTest(result, __RETRO_VERSION__, Office.context?.diagnostics?.version ?? "unknown") : "";
  const copy = async () => {
    try {
      await navigator.clipboard.writeText(report);
      setCopied("copied");
    } catch {
      setCopied("manual");
    }
  };

  return (
    <>
      <Section title="Test the shape library on your Mac">
        <Text size={200}>
          Before the library is built, this checks what your PowerPoint lets Retro copy. Select a few of the shapes you'd keep in your library (a mix is best: a styled shape with text, a group, a line, an
          icon or picture), then run the test.
        </Text>
        <Text size={200} className={ui.muted}>
          Retro rebuilds what it can beside the originals and compares the copies. Delete the copies afterwards.
        </Text>
        <div className={ui.row}>
          <Button appearance="primary" icon={<PlayRegular />} disabled={busy} onClick={() => void test()}>
            Test selected shapes
          </Button>
          <Button icon={<SlideAddRegular />} disabled={busy} onClick={() => void helper()}>
            Try the helper slide
          </Button>
        </div>
      </Section>

      {result && (result.steps.length > 0 || busy) && (
        <Section title="Results">
          <div>
            {result.steps.map((st) => (
              <div key={st.label} className={s.row}>
                <StatusIcon status={st.status} />
                <span>
                  <b>{st.label}</b>
                  <br />
                  <span className={ui.muted}>{st.detail}</span>
                </span>
              </div>
            ))}
            {busy && (
              <div className={s.row}>
                <StatusIcon status="pending" />
                <span>Working…</span>
              </div>
            )}
          </div>
          {result.preview && <img className={s.preview} src={result.preview} alt="Preview of the first selected shape" />}
          {result.items.length > 0 && (
            <div>
              {result.items.map((it, k) => (
                <div key={k} className={s.item}>
                  <span>
                    <b>{it.what}</b> <span className={ui.muted}>{it.name}</span>
                  </span>
                  {it.issues.length === 0 ? (
                    <Badge appearance="tint" color="success" style={{ alignSelf: "flex-start" }}>
                      One click
                    </Badge>
                  ) : (
                    <>
                      <Badge appearance="tint" color="warning" style={{ alignSelf: "flex-start" }}>
                        Helper slide
                      </Badge>
                      <span className={ui.muted}>Because of: {it.issues.join(", ")}</span>
                    </>
                  )}
                </div>
              ))}
            </div>
          )}
          {result.differences.length > 0 && (
            <>
              <Text weight="semibold">Differences in the copies</Text>
              <ul className={s.list}>
                {result.differences.map((d) => (
                  <li key={d}>{d}</li>
                ))}
              </ul>
            </>
          )}
          {!busy && (
            <div className={ui.row}>
              <Button icon={<CopyRegular />} onClick={() => void copy()}>
                Copy results
              </Button>
              {copied === "copied" && (
                <Text size={200} className={ui.muted} role="status">
                  Copied. Paste it into our chat.
                </Text>
              )}
            </div>
          )}
          {copied === "manual" && <textarea className={s.report} readOnly value={report} onFocus={(e) => e.currentTarget.select()} aria-label="Shape test results" />}
        </Section>
      )}
    </>
  );
}
