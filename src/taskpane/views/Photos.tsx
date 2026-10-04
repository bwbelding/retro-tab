import { useEffect, useMemo, useRef, useState } from "react";
import { Button, Input, ProgressBar, Text, makeStyles, tokens } from "@fluentui/react-components";
import { ArrowSyncRegular, ChevronRightRegular, FolderOpenRegular, FolderRegular, ImageRegular, PlugConnectedRegular, SearchRegular } from "@fluentui/react-icons";
import { activity, describeError } from "../../lib/activity";
import { isConnected, insertPhoto, photoStore, scanFolder, type ScanProgress } from "../../lib/photoActions";
import { childFolders, filterPhotos, type Anchor, type Kind, type PhotoEntry, type PhotoIndex, type Placement } from "../../lib/photos";
import { UserError } from "../../lib/ppt";
import { prefs } from "../../lib/prefs";
import { pane } from "../store";
import { Section, Seg, useUi } from "../ui";

const useStyles = makeStyles({
  grid: { display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: "6px" },
  thumb: {
    aspectRatio: "4 / 3",
    padding: 0,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusSmall,
    backgroundColor: tokens.colorNeutralBackground3,
    overflow: "hidden",
    cursor: "pointer",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    ":hover": { outline: `2px solid ${tokens.colorBrandStroke1}` },
    ":focus-visible": { outline: `2px solid ${tokens.colorStrokeFocus2}` },
  },
  img: { width: "100%", height: "100%", objectFit: "cover" },
  // Graphics are shown whole on a light checkerboard, since many are icons on transparent backgrounds.
  vector: {
    objectFit: "contain",
    padding: "6px",
    boxSizing: "border-box",
    backgroundColor: "#f3f3f3",
    backgroundImage: "linear-gradient(45deg, #e2e2e2 25%, transparent 25%, transparent 75%, #e2e2e2 75%), linear-gradient(45deg, #e2e2e2 25%, transparent 25%, transparent 75%, #e2e2e2 75%)",
    backgroundSize: "10px 10px",
    backgroundPosition: "0 0, 5px 5px",
  },
  crumbs: { display: "flex", flexWrap: "wrap", alignItems: "center", gap: "2px", fontSize: tokens.fontSizeBase200 },
  folders: { display: "flex", flexDirection: "column", maxHeight: "180px", overflowY: "auto", border: `1px solid ${tokens.colorNeutralStroke2}`, borderRadius: tokens.borderRadiusMedium },
  folder: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "5px 8px",
    border: "none",
    borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
    background: "none",
    color: tokens.colorNeutralForeground1,
    textAlign: "left",
    cursor: "pointer",
    fontSize: tokens.fontSizeBase200,
    ":hover": { backgroundColor: tokens.colorNeutralBackground1Hover },
    ":focus-visible": { outline: `2px solid ${tokens.colorStrokeFocus2}`, outlineOffset: "-2px" },
  },
  folderName: { flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  hidden: { display: "none" },
});

const PLACEMENTS: { value: Placement; label: string; help: string }[] = [
  { value: "full", label: "Full bleed", help: "Covers the whole slide, behind everything else, cropped to fit." },
  { value: "box", label: "Fill box", help: "Fills the selected shape (rectangle, circle, rounded box…), cropped to its shape." },
  { value: "background", label: "Background", help: "Replaces this slide's background, cropped to fit, and hides the template's background graphics on this slide." },
  { value: "asis", label: "As is", help: "The whole photo, uncropped, centered on the slide." },
];

const ANCHORS: { value: Anchor; label: string }[] = [
  { value: "top", label: "Top" },
  { value: "center", label: "Center" },
  { value: "bottom", label: "Bottom" },
];

const KINDS: { value: Kind; label: string }[] = [
  { value: "all", label: "All" },
  { value: "photos", label: "Photos" },
  { value: "graphics", label: "Graphics" },
];

const PAGE = 60;

function Thumb({ photo, onPick }: { photo: PhotoEntry; onPick: (p: PhotoEntry) => void }) {
  const s = useStyles();
  const [url, setUrl] = useState<string>();
  useEffect(() => {
    let made: string | undefined;
    let live = true;
    photoStore
      .thumb(photo.id)
      .then((data) => {
        if (!data || !live) return;
        made = URL.createObjectURL(new Blob([data], { type: photo.vector ? "image/svg+xml" : "image/jpeg" }));
        setUrl(made);
      })
      .catch(() => undefined);
    return () => {
      live = false;
      if (made) URL.revokeObjectURL(made);
    };
  }, [photo.id, photo.vector]);
  return (
    <button className={s.thumb} onClick={() => onPick(photo)} title={photo.id} aria-label={`Insert ${photo.name}`}>
      {url ? <img className={photo.vector ? `${s.img} ${s.vector}` : s.img} src={url} alt="" /> : <ImageRegular fontSize={20} />}
    </button>
  );
}

export function Photos() {
  const s = useStyles();
  const ui = useUi();
  const input = useRef<HTMLInputElement>(null);
  const [index, setIndex] = useState<PhotoIndex | null | undefined>();
  const [connected, setConnected] = useState(isConnected());
  const [progress, setProgress] = useState<ScanProgress | undefined>();
  const [placement, setPlacement] = useState<Placement>("full");
  const [anchor, setAnchor] = useState<Anchor>("center");
  const [query, setQuery] = useState("");
  const [folder, setFolder] = useState("");
  const [kind, setKind] = useState<Kind>("all");
  const [limit, setLimit] = useState(PAGE);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    // React doesn't know the non-standard attribute, so set it directly.
    input.current?.setAttribute("webkitdirectory", "");
    photoStore
      .index()
      .then((i) => setIndex(i ?? null))
      .catch(() => setIndex(null));
    void prefs.photo().then((p) => {
      setPlacement(p.placement);
      setAnchor(p.anchor);
    });
  }, []);

  const choose = (p: Placement, a: Anchor) => {
    setPlacement(p);
    setAnchor(a);
    void prefs.setPhoto({ placement: p, anchor: a }).catch(() => undefined);
  };

  const scan = async (files: File[]) => {
    if (files.length === 0) return;
    pane.showMessage(undefined);
    try {
      const r = await activity.track("pane photo scan", () => scanFolder(files, setProgress));
      setIndex(r.index);
      setConnected(true);
      setFolder("");
      const graphics = r.index.photos.filter((p) => p.vector).length;
      const parts = [`${r.index.photos.length - graphics} photos${graphics ? ` and ${graphics} graphics` : ""} in “${r.index.root}”`];
      if (r.added) parts.push(`${r.added} new`);
      if (r.removed) parts.push(`${r.removed} removed`);
      if (r.unreadable) parts.push(`${r.unreadable} couldn't be read (formats PowerPoint can't open, or cloud files not downloaded to this Mac)`);
      const st = r.stats;
      parts.push(`Finder gave Retro ${st.files} files in ${st.folders} folder${st.folders === 1 ? "" : "s"}, ${st.depth} level${st.depth === 1 ? "" : "s"} of subfolders deep`);
      if (st.skipped.length) parts.push(`skipped: ${st.skipped.slice(0, 6).map((x) => `${x.count} ${x.ext}`).join(", ")}${st.skipped.length > 6 ? ", …" : ""}`);
      pane.showMessage({ intent: r.unreadable ? "warning" : "success", text: `${parts.join(" · ")}.` });
    } catch (e) {
      pane.showMessage({ intent: e instanceof UserError ? "warning" : "error", text: e instanceof UserError ? e.message : describeError(e) });
    } finally {
      setProgress(undefined);
    }
  };

  const pick = async (photo: PhotoEntry) => {
    if (busy) return;
    setBusy(true);
    pane.showMessage(undefined);
    try {
      const text = await activity.track("pane insert photo", () => insertPhoto(photo, placement, anchor));
      pane.showMessage({ intent: "success", text });
    } catch (e) {
      pane.showMessage({ intent: e instanceof UserError ? "warning" : "error", text: e instanceof UserError ? e.message : describeError(e) });
    } finally {
      setBusy(false);
    }
  };

  const folders = useMemo(() => childFolders(index?.photos ?? [], folder), [index, folder]);
  const shown = useMemo(() => filterPhotos(index?.photos ?? [], query, folder, kind), [index, query, folder, kind]);
  const graphicsCount = useMemo(() => (index?.photos ?? []).filter((p) => p.vector).length, [index]);
  const hasGraphics = graphicsCount > 0;
  const openFolder = (path: string) => {
    setFolder(path);
    setLimit(PAGE);
  };
  const crumbs = folder ? folder.split("/") : [];
  const help = PLACEMENTS.find((p) => p.value === placement)!.help;

  return (
    <>
      <input
        ref={input}
        className={s.hidden}
        type="file"
        multiple
        aria-label="Choose your photo folder"
        onChange={(e) => {
          void scan(Array.from(e.currentTarget.files ?? []));
          e.currentTarget.value = "";
        }}
      />

      <Section title="Photos">
        {index === undefined ? (
          <Text size={200} className={ui.muted}>
            Loading…
          </Text>
        ) : index === null ? (
          <>
            <Text size={200}>
              Choose the folder that holds your photos. OneDrive and Box folders work too: pick them from the sidebar in the Finder window. Subfolders become filters.
            </Text>
            <Button appearance="primary" icon={<FolderOpenRegular />} style={{ alignSelf: "flex-start" }} onClick={() => input.current?.click()} disabled={Boolean(progress)}>
              Choose photo folder…
            </Button>
          </>
        ) : (
          <>
            <Text size={200}>
              <b>{index.root}</b> · {index.photos.length - graphicsCount} photos{graphicsCount ? ` · ${graphicsCount} graphics` : ""}
            </Text>
            {!connected && (
              <Text size={200} className={ui.muted}>
                Previews are ready. To insert, click Reconnect and choose the same folder in Finder (PowerPoint forgets folder access when it restarts).
              </Text>
            )}
            <div className={ui.row}>
              {connected ? (
                <Button icon={<ArrowSyncRegular />} onClick={() => input.current?.click()} disabled={Boolean(progress)}>
                  Rescan
                </Button>
              ) : (
                <Button appearance="primary" icon={<PlugConnectedRegular />} onClick={() => input.current?.click()} disabled={Boolean(progress)}>
                  Reconnect
                </Button>
              )}
              <Button icon={<FolderOpenRegular />} onClick={() => input.current?.click()} disabled={Boolean(progress)}>
                Change folder…
              </Button>
            </div>
          </>
        )}
        {progress && (
          <>
            <ProgressBar value={progress.total ? progress.done / progress.total : undefined} />
            <Text size={200} className={ui.muted}>
              {progress.total ? `Making previews: ${progress.done} of ${progress.total}. This is quicker next time.` : "Reading the folder…"}
            </Text>
          </>
        )}
      </Section>

      {index && index.photos.length > 0 && (
        <>
          <Section>
            <Seg label="Placement" value={placement} options={PLACEMENTS} onChange={(p) => choose(p, anchor)} />
            {placement !== "asis" && <Seg label="Crop from" value={anchor} options={ANCHORS} onChange={(a) => choose(placement, a)} />}
            <Text size={200} className={ui.muted}>
              {help}
              {placement !== "asis" && " Tall photos are trimmed from the top, center or bottom as chosen."}
              {hasGraphics && " Graphics (SVG) always go in whole: inside the selected shape, or centered on the slide."}
            </Text>
          </Section>

          <Section>
            <Input contentBefore={<SearchRegular />} placeholder="Search file names" value={query} onChange={(_, d) => setQuery(d.value)} aria-label="Search photos" />
            {hasGraphics && (
              <Seg
                label="Show"
                value={kind}
                options={KINDS}
                onChange={(k) => {
                  setKind(k);
                  setLimit(PAGE);
                }}
              />
            )}
            <nav className={s.crumbs} aria-label="Folder">
              {[index.root, ...crumbs].map((name, i) => {
                const path = crumbs.slice(0, i).join("/");
                const here = i === crumbs.length;
                return (
                  <span key={i} style={{ display: "contents" }}>
                    {i > 0 && <ChevronRightRegular fontSize={12} aria-hidden />}
                    {here ? (
                      <Text size={200} weight="semibold" style={{ padding: "0 6px" }} aria-current="location">
                        {name}
                      </Text>
                    ) : (
                      <Button size="small" appearance="outline" onClick={() => openFolder(path)}>
                        {name}
                      </Button>
                    )}
                  </span>
                );
              })}
            </nav>
            {folders.length > 0 && (
              <div className={s.folders}>
                {folders.map((f) => (
                  <button key={f.path} className={s.folder} onClick={() => openFolder(f.path)}>
                    <FolderRegular fontSize={16} aria-hidden />
                    <span className={s.folderName}>{f.name}</span>
                    <span className={ui.muted}>{f.count}</span>
                  </button>
                ))}
              </div>
            )}
            {shown.length === 0 ? (
              <Text size={200} className={ui.muted}>
                Nothing matches.
              </Text>
            ) : (
              <>
                <div className={s.grid}>
                  {shown.slice(0, limit).map((p) => (
                    <Thumb key={p.id} photo={p} onPick={(photo) => void pick(photo)} />
                  ))}
                </div>
                <Text size={200} className={ui.muted}>
                  Showing {Math.min(limit, shown.length)} of {shown.length}
                  {folders.length > 0 ? " (including subfolders)" : ""}. Click one to place it.
                </Text>
                {shown.length > limit && (
                  <Button size="small" style={{ alignSelf: "flex-start" }} onClick={() => setLimit((n) => n + PAGE)}>
                    Show more
                  </Button>
                )}
              </>
            )}
          </Section>
        </>
      )}
    </>
  );
}
