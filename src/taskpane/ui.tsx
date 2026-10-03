// Small layout pieces shared by the pane's views, matching the design canvas.

import type { ReactNode } from "react";
import { Text, ToggleButton, makeStyles, tokens } from "@fluentui/react-components";

const useStyles = makeStyles({
  section: {
    padding: "14px 16px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    display: "flex",
    flexDirection: "column",
    gap: "10px",
  },
  heading: { margin: 0 },
  seg: { display: "flex", gap: "4px", flexWrap: "wrap" },
  segItem: { flex: "1 1 0", minWidth: "fit-content" },
  row: { display: "flex", gap: "8px", alignItems: "flex-end", flexWrap: "wrap" },
  grid2: { display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: "8px" },
  grid3: { display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: "6px" },
  muted: { color: tokens.colorNeutralForeground3 },
});

export const useUi = useStyles;

export function Section({ title, children }: { title?: string; children: ReactNode }) {
  const s = useStyles();
  return (
    <section className={s.section}>
      {title && (
        <Text as="h2" weight="semibold" className={s.heading}>
          {title}
        </Text>
      )}
      {children}
    </section>
  );
}

/** A segmented choice: one toggle per option, exactly one selected. */
export function Seg<T extends string>({
  label,
  value,
  options,
  onChange,
}: {
  label: string;
  value: T;
  options: ReadonlyArray<{ value: T; label: string }>;
  onChange: (v: T) => void;
}) {
  const s = useStyles();
  return (
    <div className={s.seg} role="radiogroup" aria-label={label}>
      {options.map((o) => (
        <ToggleButton
          key={o.value}
          size="small"
          className={s.segItem}
          role="radio"
          aria-checked={o.value === value}
          checked={o.value === value}
          onClick={() => onChange(o.value)}
        >
          {o.label}
        </ToggleButton>
      ))}
    </div>
  );
}
