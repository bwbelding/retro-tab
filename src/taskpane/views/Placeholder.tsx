import { Text, makeStyles, tokens } from "@fluentui/react-components";
import type { ViewInfo } from "../views";

const useStyles = makeStyles({
  root: { padding: "32px 20px", display: "flex", flexDirection: "column", alignItems: "center", gap: "10px", textAlign: "center" },
  icon: { color: tokens.colorNeutralForeground3 },
  heading: { margin: 0 },
});

export function Placeholder({ view }: { view: ViewInfo }) {
  const styles = useStyles();
  const Icon = view.icon;
  return (
    <section className={styles.root}>
      <Icon fontSize={48} className={styles.icon} aria-hidden />
      <Text as="h2" size={400} weight="semibold" className={styles.heading}>
        {view.title}
      </Text>
      <Text>{view.coming}</Text>
    </section>
  );
}
