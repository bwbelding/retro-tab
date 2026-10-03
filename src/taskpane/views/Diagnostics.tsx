import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Button, MessageBar, MessageBarBody, Spinner, Text, makeStyles, tokens } from "@fluentui/react-components";
import {
  ArrowClockwiseRegular,
  ArrowDownloadRegular,
  ArrowUploadRegular,
  CheckmarkCircleRegular,
  CopyRegular,
  DismissCircleRegular,
  FolderOpenRegular,
  WarningRegular,
} from "@fluentui/react-icons";
import { activity as activityLog, describeError, formatActivity, type ActivityEntry } from "../../lib/activity";
import type { Backup } from "../../lib/library";
import { library } from "../../lib/libraryActions";
import { kv } from "../../lib/idb";
import { officeIsSetSupported, probeRequirements } from "../../lib/capabilities";
import {
  checkHosting,
  checkPowerPoint,
  checkSharedStorage,
  checkStorage,
  formatReport,
  interpretFolderPick,
  summarize,
  type Check,
  type Environment,
  type Status,
} from "../../lib/diagnostics";

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", minHeight: "100%" },
  section: {
    padding: "14px 16px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    display: "flex",
    flexDirection: "column",
    gap: "10px",
  },
  kv: { display: "grid", gridTemplateColumns: "96px minmax(0, 1fr)", gap: "6px 12px", margin: 0, fontSize: tokens.fontSizeBase200 },
  dt: { color: tokens.colorNeutralForeground3 },
  dd: { margin: 0, fontWeight: tokens.fontWeightSemibold, overflowWrap: "anywhere" },
  row: {
    display: "grid",
    gridTemplateColumns: "18px minmax(0, 1fr)",
    gap: "8px",
    padding: "8px 0",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    fontSize: tokens.fontSizeBase200,
    lineHeight: tokens.lineHeightBase200,
    ":first-child": { borderTop: "none" },
  },
  muted: { color: tokens.colorNeutralForeground3 },
  ok: { color: tokens.colorPaletteGreenForeground1 },
  warn: { color: tokens.colorPaletteMarigoldForeground1 },
  fail: { color: tokens.colorPaletteRedForeground1 },
  footer: {
    position: "sticky",
    bottom: 0,
    marginTop: "auto",
    display: "flex",
    gap: "8px",
    alignItems: "center",
    padding: "10px 16px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  report: { width: "100%", height: "120px", fontFamily: "ui-monospace, Menlo, monospace", fontSize: "11px" },
  hidden: { display: "none" },
  heading: { margin: 0 },
});

const ICON = { ok: CheckmarkCircleRegular, warn: WarningRegular, fail: DismissCircleRegular } as const;
const LABEL = { ok: "OK", warn: "Warning", fail: "Failed" } as const;

export function StatusIcon({ status }: { status: Status }) {
  const styles = useStyles();
  if (status === "pending") return <Spinner size="extra-tiny" aria-label="Checking" />;
  const Icon = ICON[status];
  return <Icon fontSize={16} className={styles[status]} aria-label={LABEL[status]} />;
}

function environment(tab: string): Environment {
  const d = typeof Office !== "undefined" ? Office.context?.diagnostics : undefined;
  return {
    host: d?.host ? String(d.host) : "not in PowerPoint",
    platform: d?.platform ? String(d.platform) : "browser",
    officeVersion: d?.version ?? "unknown",
    tab: tab === "polish" ? "Retro Polish" : "Retro Build",
    retroVersion: __RETRO_VERSION__,
    origin: location.origin,
    userAgent: navigator.userAgent,
  };
}

function downloadJson(name: string, value: unknown) {
  const url = URL.createObjectURL(new Blob([JSON.stringify(value)], { type: "application/json" }));
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

/** Back up the shape library and settings to a file, or restore from one. */
function BackupSection() {
  const styles = useStyles();
  const input = useRef<HTMLInputElement>(null);
  const [status, setStatus] = useState<{ intent: "success" | "error" | "info"; text: string } | undefined>();

  const backup = async () => {
    try {
      const data = await activityLog.track("pane backup", () => library.backup());
      downloadJson(`Retro backup ${data.exported.slice(0, 10)}.json`, data);
      setStatus({
        intent: "info",
        text: `Backup of ${data.items.length} shape${data.items.length === 1 ? "" : "s"} and your settings saved to Downloads. If no file appeared there, tell me: PowerPoint may block downloads from add-ins.`,
      });
    } catch (e) {
      setStatus({ intent: "error", text: describeError(e) });
    }
  };

  const restore = async (file: File | undefined) => {
    if (!file) return;
    try {
      const added = await activityLog.track("pane restore", async () => library.restore(JSON.parse(await file.text()) as Backup));
      setStatus({ intent: "success", text: `Restored ${added} shape${added === 1 ? "" : "s"}${added ? "" : " (they were all already in your library)"} and your settings.` });
    } catch (e) {
      setStatus({ intent: "error", text: e instanceof SyntaxError ? "That file isn't a Retro backup." : describeError(e) });
    }
  };

  return (
    <div className={styles.section}>
      <Text as="h3" weight="semibold" className={styles.heading}>
        Backup
      </Text>
      <Text size={200} className={styles.muted}>
        Your shape library and settings live in PowerPoint's add-in storage on this Mac. Clearing Office's cache would erase them, so keep a backup file.
      </Text>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
        <Button icon={<ArrowDownloadRegular />} onClick={() => void backup()}>
          Back up
        </Button>
        <Button icon={<ArrowUploadRegular />} onClick={() => input.current?.click()}>
          Restore…
        </Button>
      </div>
      <input
        ref={input}
        className={styles.hidden}
        type="file"
        accept=".json,application/json"
        aria-label="Choose a Retro backup file"
        onChange={(e) => {
          void restore(e.currentTarget.files?.[0]);
          e.currentTarget.value = "";
        }}
      />
      {status && (
        <MessageBar layout="multiline" intent={status.intent}>
          <MessageBarBody>{status.text}</MessageBarBody>
        </MessageBar>
      )}
    </div>
  );
}

export function Diagnostics({ tab }: { tab: "build" | "polish" }) {
  const styles = useStyles();
  const env = useMemo(() => environment(tab), [tab]);
  const requirements = useMemo(() => probeRequirements(officeIsSetSupported()), []);
  const [checks, setChecks] = useState<Check[]>([]);
  const [folder, setFolder] = useState<Check>({
    id: "folder",
    label: "Photo folder picking",
    status: "warn",
    detail: "Not tested yet. Click “Test folder picking” and choose any folder with photos in it.",
  });
  const [copied, setCopied] = useState<"" | "copied" | "manual">("");
  const [recent, setRecent] = useState<ActivityEntry[]>([]);
  const folderInput = useRef<HTMLInputElement>(null);

  const run = useCallback(async () => {
    setChecks([]);
    const other = tab === "build" ? "polish" : "build";
    const results = await Promise.all([checkPowerPoint(), checkStorage(kv), checkSharedStorage(kv, tab, other), checkHosting()]);
    setChecks(results);
    setRecent(await activityLog.read());
  }, [tab]);

  useEffect(() => {
    // Run the checks once when the pane opens; results arrive asynchronously.
    // eslint-disable-next-line react-hooks/set-state-in-effect
    void run();
  }, [run]);

  useEffect(() => {
    // React doesn't know the non-standard attribute, so set it directly.
    folderInput.current?.setAttribute("webkitdirectory", "");
  }, []);

  const allChecks = checks.length ? [...checks, folder] : [];
  const summary = summarize(allChecks.length ? allChecks : [{ id: "x", label: "", status: "pending", detail: "" }], requirements);
  const report = formatReport(env, requirements, allChecks, new Date(), formatActivity(recent));

  const copy = async () => {
    try {
      await navigator.clipboard.writeText(report);
      setCopied("copied");
    } catch {
      setCopied("manual"); // Clipboard blocked: show the text so it can be copied by hand.
    }
  };

  return (
    <section className={styles.root}>
      <div className={styles.section}>
        <Text as="h2" size={400} weight="semibold" className={styles.heading}>
          Diagnostics
        </Text>
        <MessageBar layout="multiline" intent={summary.status === "ok" ? "success" : summary.status === "fail" ? "error" : summary.status === "warn" ? "warning" : "info"}>
          <MessageBarBody>{summary.text}</MessageBarBody>
        </MessageBar>
      </div>

      <div className={styles.section}>
        <Text as="h3" weight="semibold" className={styles.heading}>
          This Mac
        </Text>
        <dl className={styles.kv}>
          <dt className={styles.dt}>PowerPoint</dt>
          <dd className={styles.dd}>{env.officeVersion}</dd>
          <dt className={styles.dt}>Platform</dt>
          <dd className={styles.dd}>{env.platform}</dd>
          <dt className={styles.dt}>Retro</dt>
          <dd className={styles.dd}>
            v{env.retroVersion} · {env.tab} tab
          </dd>
          <dt className={styles.dt}>Loaded from</dt>
          <dd className={styles.dd}>{location.host}</dd>
        </dl>
      </div>

      <div className={styles.section}>
        <Text as="h3" weight="semibold" className={styles.heading}>
          What PowerPoint allows
        </Text>
        <div>
          {requirements.map((r) => (
            <div key={`${r.set}${r.version}`} className={styles.row}>
              <StatusIcon status={r.supported ? "ok" : "fail"} />
              <span>
                <b>{r.label}</b>
                <br />
                <span className={styles.muted}>
                  {r.set} {r.version} · {r.supported ? r.unlocks : r.macMin ? `needs PowerPoint ${r.macMin} or later` : "not available here"}
                </span>
              </span>
            </div>
          ))}
        </div>
      </div>

      <div className={styles.section}>
        <Text as="h3" weight="semibold" className={styles.heading}>
          Your company's setup
        </Text>
        <div>
          {checks.length === 0 && (
            <div className={styles.row}>
              <StatusIcon status="pending" />
              <span>Running checks…</span>
            </div>
          )}
          {allChecks.map((c) => (
            <div key={c.id} className={styles.row}>
              <StatusIcon status={c.status} />
              <span>
                <b>{c.label}</b>
                <br />
                <span className={styles.muted}>{c.detail}</span>
              </span>
            </div>
          ))}
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <Button icon={<FolderOpenRegular />} onClick={() => folderInput.current?.click()}>
            Test folder picking
          </Button>
          <Button appearance="subtle" icon={<ArrowClockwiseRegular />} onClick={() => void run()}>
            Run again
          </Button>
        </div>
        <input
          ref={folderInput}
          className={styles.hidden}
          type="file"
          multiple
          aria-label="Choose a photo folder"
          onChange={(e) => setFolder(interpretFolderPick(Array.from(e.currentTarget.files ?? [])))}
        />
      </div>

      <BackupSection />

      <div className={styles.section}>
        <Text as="h3" weight="semibold" className={styles.heading}>
          Recent Retro actions
        </Text>
        {recent.length === 0 ? (
          <Text size={200} className={styles.muted}>
            None yet. Buttons you press are listed here, so problems show up in the report.
          </Text>
        ) : (
          <div>
            {[...recent]
              .reverse()
              .slice(0, 10)
              .map((e, i) => (
                <div key={`${e.at}${i}`} className={styles.row}>
                  <StatusIcon status={e.status === "ok" ? "ok" : e.status === "error" ? "fail" : "warn"} />
                  <span>
                    <b>{e.action}</b> <span className={styles.muted}>{new Date(e.at).toLocaleTimeString()}{e.ms !== undefined ? ` · ${e.ms} ms` : ""}</span>
                    {e.status === "started" && <span className={styles.muted}> · started, no finish recorded yet</span>}
                    {e.detail && (
                      <>
                        <br />
                        <span className={styles.muted}>{e.detail}</span>
                      </>
                    )}
                  </span>
                </div>
              ))}
          </div>
        )}
      </div>

      {copied === "manual" && (
        <div className={styles.section}>
          <Text size={200}>Copying is blocked here. Select all of this text, then press ⌘C:</Text>
          <textarea className={styles.report} readOnly value={report} onFocus={(e) => e.currentTarget.select()} aria-label="Diagnostics report" />
        </div>
      )}

      <div className={styles.footer}>
        <Button appearance="primary" icon={<CopyRegular />} onClick={() => void copy()} disabled={!checks.length}>
          Copy report
        </Button>
        {copied === "copied" && (
          <Text size={200} className={styles.muted} role="status">
            Copied. Paste it into our chat.
          </Text>
        )}
      </div>
    </section>
  );
}
