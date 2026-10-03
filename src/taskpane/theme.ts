import { createDarkTheme, createLightTheme, type BrandVariants, type Theme } from "@fluentui/react-components";

// Fluent brand ramp built around PowerPoint's accent (#C43E1C at step 80).
const retro: BrandVariants = {
  10: "#170804", 20: "#2E0F09", 30: "#46170D", 40: "#5E1E11", 50: "#772614", 60: "#902D17",
  70: "#AA351A", 80: "#C43E1C", 90: "#CD5131", 100: "#D56446", 110: "#DC775B", 120: "#E38A71",
  130: "#E99D87", 140: "#EFB09E", 150: "#F4C3B5", 160: "#F9D7CD",
};

export const lightTheme: Theme = createLightTheme(retro);
export const darkTheme: Theme = { ...createDarkTheme(retro), colorBrandForeground1: retro[110], colorBrandForegroundLink: retro[110] };

/** Follow PowerPoint's theme when Office reports it, otherwise the Mac's appearance. */
export function prefersDark(): boolean {
  const bg = typeof Office !== "undefined" ? Office.context?.officeTheme?.bodyBackgroundColor : undefined;
  if (bg && /^#?[0-9a-f]{6}$/i.test(bg)) {
    const n = parseInt(bg.replace("#", ""), 16);
    const luminance = (0.299 * (n >> 16) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255)) / 255;
    return luminance < 0.5;
  }
  return window.matchMedia?.("(prefers-color-scheme: dark)").matches ?? false;
}
