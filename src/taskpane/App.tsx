import { useEffect, useState } from "react";
import { FluentProvider, makeStyles, mergeClasses, tokens } from "@fluentui/react-components";
import { darkTheme, lightTheme, prefersDark } from "./theme";
import { VIEWS, type ViewKey } from "./views";
import { Diagnostics } from "./views/Diagnostics";
import { Placeholder } from "./views/Placeholder";

const useStyles = makeStyles({
  root: {
    height: "100vh",
    display: "flex",
    flexDirection: "column",
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
  },
  nav: {
    display: "flex",
    flexShrink: 0,
    padding: "0 2px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  navItem: {
    flex: "1 1 0",
    minWidth: 0,
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "2px",
    padding: "8px 0 6px",
    border: "none",
    borderBottom: "2px solid transparent",
    background: "transparent",
    color: tokens.colorNeutralForeground3,
    fontFamily: "inherit",
    fontSize: tokens.fontSizeBase100,
    lineHeight: "14px",
    cursor: "pointer",
    ":hover": { color: tokens.colorNeutralForeground1 },
    ":focus-visible": { outline: `2px solid ${tokens.colorStrokeFocus2}`, outlineOffset: "-2px" },
  },
  navOn: {
    color: tokens.colorNeutralForeground1,
    fontWeight: tokens.fontWeightSemibold,
    borderBottomColor: tokens.colorBrandStroke1,
  },
  body: { flex: 1, minHeight: 0, overflowY: "auto" },
});

export function App({ initialView, tab }: { initialView: ViewKey; tab: "build" | "polish" }) {
  const styles = useStyles();
  const [view, setView] = useState<ViewKey>(initialView);
  const [dark, setDark] = useState(prefersDark);

  useEffect(() => {
    const media = window.matchMedia?.("(prefers-color-scheme: dark)");
    const update = () => setDark(prefersDark());
    media?.addEventListener("change", update);
    return () => media?.removeEventListener("change", update);
  }, []);

  const go = (key: ViewKey) => {
    setView(key);
    const url = new URL(location.href);
    url.searchParams.set("view", key);
    history.replaceState(null, "", url);
  };

  const current = VIEWS.find((v) => v.key === view)!;

  return (
    <FluentProvider theme={dark ? darkTheme : lightTheme} className={styles.root}>
      <nav className={styles.nav} aria-label="Retro sections">
        {VIEWS.map(({ key, nav, title, icon: Icon }) => (
          <button
            key={key}
            className={mergeClasses(styles.navItem, key === view && styles.navOn)}
            aria-current={key === view ? "page" : undefined}
            aria-label={key === "diagnostics" ? "More: Diagnostics and Backup" : undefined}
            title={title}
            onClick={() => go(key)}
          >
            <Icon fontSize={20} aria-hidden />
            {nav}
          </button>
        ))}
      </nav>
      <main className={styles.body}>
        {view === "diagnostics" ? <Diagnostics tab={tab} /> : <Placeholder view={current} />}
      </main>
    </FluentProvider>
  );
}
