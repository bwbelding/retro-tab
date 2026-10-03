import { useEffect, useMemo, useRef, useState } from "react";
import { Button, Input, ProgressBar, Text, ToggleButton, makeStyles, tokens } from "@fluentui/react-components";
import { ArrowSyncRegular, FolderOpenRegular, ImageRegular, PlugConnectedRegular, SearchRegular } from "@fluentui/react-icons";
import { activity, describeError } from "../../lib/activity";
import { isConnected, insertPhoto, photoStore, scanFolder, type ScanProgress } from "../../lib/photoActions";
import { filterPhotos, topFolders, type Anchor, type PhotoEntry, type PhotoIndex, type Placement } from "../../lib/photos";
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
  chips: { display: "flex", gap: "4px", flexWrap: "wrap" },
  hidden: { display: "none" },
});

const PLACEMENTS: { value: Placement; label: string; help: string }[] = [
  { value: "full", label: "Full bleed", help: "Covers the whole slide, behind everything else, cropped to fit." },
  { value: "box", label: "Fill box", help: "Fills the selected shape (rectangle, circle, rounded box…), cropped to its shape." },
  { value: "background", label: "Background", help: "Becomes this slide's background, cropped to fit." },
  { value: "asis", label: "As is", help: "The whole photo, uncropped, centered on the slide." },
];

const ANCHORS: { value: Anchor; label: string }[] = [
  { value: "top", label: "Top" },
  { value: "center", label: "Center" },
  { value: "bottom", label: "Bottom" },
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
        made = URL.createObjectURL(new Blob([data], { type: "image/jpeg" }));
        setUrl(made);
      })
      .catch(() => undefined);
    return () => {
      live = false;
      if (made) URL.revokeObjectURL(made);
    };
  }, [photo.id]);
  return (
    <button className={s.thumb} onClick={() => onPick(photo)} title={photo.id} aria-label={`Insert ${photo.name}`}>
      {url ? <img className={s.img} src={url} alt="" /> : <ImageRegular fontSize={20} />}
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
      const parts = [`${r.index.photos.length} photos in “${r.index.root}”`];
      if (r.added) parts.push(`${r.added} new`);
      if (r.removed) parts.push(`${r.removed} removed`);
      if (r.unreadable) parts.push(`${r.unreadable} couldn't be read (formats PowerPoint can't open, or cloud files not downloaded to this Mac)`);
      const st = r.stats;
      parts.push(`Finder gave Retro ${st.files} files in ${st.folders} folder${st.folders === 1 ? "" : "s"}, ${st.depth} level${st.depth === 1 ? "" : "s"} of subfolders deep`);
      if (st.skipped.length) parts.push(`not photos: ${st.skipped.slice(0, 6).map((x) => `${x.count} ${x.ext}`).join(", ")}${st.skipped.length > 6 ? ", …" : ""}`);
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

  const folders = useMemo(() => topFolders(index?.photos ?? []), [index]);
  const shown = useMemo(() => filterPhotos(index?.photos ?? [], query, folder), [index, query, folder]);
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
              <b>{index.root}</b> · {index.photos.length} photos
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
              <Button appearance="subtle" icon={<FolderOpenRegular />} onClick={() => input.current?.click()} disabled={Boolean(progress)}>
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
            </Text>
          </Section>

          <Section>
            <Input contentBefore={<SearchRegular />} placeholder="Search file names" value={query} onChange={(_, d) => setQuery(d.value)} aria-label="Search photos" />
            {folders.length > 0 && (
              <div className={s.chips} role="radiogroup" aria-label="Folder">
                {["", ...folders].map((f) => (
                  <ToggleButton
                    key={f || "all"}
                    size="small"
                    role="radio"
                    aria-checked={f === folder}
                    checked={f === folder}
                    onClick={() => {
                      setFolder(f);
                      setLimit(PAGE);
                    }}
                  >
                    {f || "All"}
                  </ToggleButton>
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
                  Showing {Math.min(limit, shown.length)} of {shown.length}. Click a photo to place it.
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
