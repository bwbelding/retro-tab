// Ribbon icon artwork: inner SVG markup on a 24x24 grid. scripts/build-icons.mjs renders each
// one to PNG at the sizes Office asks for. `currentColor` becomes the Retro accent, which stays
// readable on both the light and dark PowerPoint ribbons (Office does not recolor add-in icons).

export const ACCENT = "#C43E1C";

const stroke = (body) =>
  `<g fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">${body}</g>`;
const swatch = (color) =>
  `<rect x="3.5" y="3.5" width="17" height="17" rx="2.5" fill="${color}" stroke="#00000033"/>`;

export const ICONS = {
  retro: `<rect x="1" y="1" width="22" height="22" rx="5" fill="currentColor"/><g fill="#FFFFFF"><rect x="5" y="6" width="14" height="3" rx="1"/><rect x="5" y="10.5" width="9" height="3" rx="1"/><rect x="5" y="15" width="11" height="3" rx="1"/></g>`,
  sticky: stroke(`<path d="M4 4h16v11l-5 5H4z"/><path d="M15 20v-5h5"/>`),
  stamp: stroke(`<g transform="rotate(-8 12 12)"><rect x="3" y="7" width="18" height="10" rx="1"/><path d="M8 12.5h8"/></g>`),
  removeNotes: stroke(`<path d="M4 15l9-9 7 7-6 6H8z"/><path d="M9 20h11"/>`),
  library: stroke(`<rect x="3" y="3" width="7.5" height="7.5" rx="1"/><circle cx="17.25" cy="6.75" r="3.75"/><path d="M6.75 13.5l3.75 7.5H3z"/><rect x="13.5" y="13.5" width="7.5" height="7.5" rx="3.75"/>`),
  saveShape: stroke(`<rect x="3" y="3" width="12" height="12" rx="1" stroke-dasharray="3 2"/><path d="M18 13v8M14 17h8"/>`),
  recent: stroke(`<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>`),
  findUpdate: stroke(`<circle cx="10" cy="10" r="6"/><path d="M14.5 14.5L20 20"/><path d="M7.5 10h5M10 7.5v5"/>`),
  slides: stroke(`<rect x="7" y="3" width="14" height="10" rx="1"/><path d="M4 7v9a1 1 0 0 0 1 1h12"/><path d="M3 20h12"/>`),
  photos: stroke(`<rect x="3" y="4" width="18" height="16" rx="1.5"/><circle cx="9" cy="9.5" r="1.8"/><path d="M3 17l5-5 4 4 3-3 6 6"/>`),
  text: stroke(`<path d="M4 6V4h16v2M12 4v16M9 20h6"/>`),
  arrange: stroke(`<path d="M4 3v18"/><rect x="7" y="6" width="12" height="4" rx="1"/><rect x="7" y="14" width="8" height="4" rx="1"/>`),
  matchWidth: stroke(`<path d="M3 12h18M6 9l-3 3 3 3M18 9l3 3-3 3M3 5v14M21 5v14"/>`),
  matchHeight: stroke(`<path d="M12 3v18M9 6l3-3 3 3M9 18l3 3 3-3M5 3h14M5 21h14"/>`),
  matchSize: stroke(`<rect x="4" y="4" width="16" height="16" rx="1"/><path d="M9 15l6-6M11 9h4v4"/>`),
  swap: stroke(`<path d="M4 8h15l-3-3M20 16H5l3 3"/>`),
  gapDistribute: stroke(`<rect x="3" y="6" width="5" height="12" rx="1"/><rect x="16" y="6" width="5" height="12" rx="1"/><path d="M10 12h4M10 10v4M14 10v4"/>`),
  exactSize: stroke(`<path d="M3 8h18M3 16h18M8 3v18M16 3v18"/>`),
  format: stroke(`<rect x="4" y="3" width="14" height="6" rx="1"/><path d="M18 6h2v5h-8v3"/><rect x="10.5" y="14" width="3" height="7" rx="1"/>`),
  brandKit: stroke(`<path d="M12 3a9 9 0 1 0 0 18c1.2 0 2-.8 2-2 0-1.4-1.2-1.6-1.2-2.8 0-1 .8-1.7 1.8-1.7H17a4 4 0 0 0 4-4c0-4.2-4-7.5-9-7.5z"/><circle cx="7.5" cy="11" r="1"/><circle cx="10" cy="7" r="1"/><circle cx="14.5" cy="7" r="1"/>`),
  fill1: swatch("#1B4F8C"),
  fill2: swatch("#E07A1F"),
  fill3: swatch("#2E8B6E"),
  applyFonts: stroke(`<path d="M3 19l5-14 5 14M5 14h6M15 12.5a3 3 0 1 1 0 4.5M18.5 11v8"/>`),
  deckCheck: stroke(`<path d="M12 3l8 3v6c0 4.5-3.4 8-8 9-4.6-1-8-4.5-8-9V6z"/><path d="M8.5 12l2.5 2.5 4.5-5"/>`),
  tracker: stroke(`<rect x="3" y="5" width="18" height="4" rx="1"/><rect x="3" y="5" width="7" height="4" rx="1" fill="currentColor"/><path d="M3 14h18M3 18h12"/>`),
  diagnostics: stroke(`<path d="M3 12h4l2-6 4 12 2-6h6"/>`),
  backup: stroke(`<path d="M12 3v12M7 10l5 5 5-5M4 17v3h16v-3"/>`),
  smart: stroke(`<circle cx="7.5" cy="7.5" r="4.5"/><path d="M7.5 3a4.5 4.5 0 0 1 4.5 4.5H7.5z" fill="currentColor"/><path d="M17 3.5l1.3 2.7 3 .4-2.2 2.1.5 3-2.6-1.4-2.6 1.4.5-3-2.2-2.1 3-.4z"/><rect x="3" y="15" width="18" height="5" rx="2.5"/><rect x="3" y="15" width="10" height="5" rx="2.5" fill="currentColor"/>`),
  smartNumber: `<circle cx="12" cy="12" r="10" fill="currentColor"/><path d="M10 8.5l2.5-1.5v10" fill="none" stroke="#FFFFFF" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>`,
  smartNumberSet: `<circle cx="5" cy="12" r="4" fill="currentColor"/><circle cx="12" cy="12" r="4" fill="currentColor"/><circle cx="19" cy="12" r="4" fill="currentColor"/><path d="M5 10.3v3.4M11 10.6h1.6v2.8M17.7 10.6h1.8l-1.6 2.9" fill="none" stroke="#FFFFFF" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/>`,
  smartHarveySet: stroke(`<circle cx="5" cy="12" r="3.6"/><circle cx="12" cy="12" r="3.6"/><path d="M12 8.4a3.6 3.6 0 0 1 0 7.2z" fill="currentColor"/><circle cx="19" cy="12" r="3.6" fill="currentColor"/>`),
  smartHarvey: stroke(`<circle cx="12" cy="12" r="9"/><path d="M12 3a9 9 0 0 1 9 9H12z" fill="currentColor"/>`),
  smartTraffic: `<circle cx="12" cy="12" r="9" fill="#1E8E3E"/>`,
  smartTraffic3: `<rect x="7" y="2" width="10" height="20" rx="3" fill="#333F48"/><circle cx="12" cy="6.5" r="2.4" fill="#D93025"/><circle cx="12" cy="12" r="2.4" fill="#F2A900"/><circle cx="12" cy="17.5" r="2.4" fill="#1E8E3E"/>`,
  smartProgress: stroke(`<rect x="2" y="9" width="20" height="6" rx="3"/><rect x="2" y="9" width="12" height="6" rx="3" fill="currentColor"/>`),
  smartStars: stroke(`<path d="M12 3l2.7 5.6 6.1.9-4.4 4.3 1 6.1L12 17l-5.4 2.9 1-6.1L3.2 9.5l6.1-.9z" fill="#F2B400" stroke="#F2B400"/>`),
  smartCheckbox: stroke(`<rect x="3.5" y="3.5" width="17" height="17" rx="3"/><path d="M7.5 12.5l3 3 6-7"/>`),
  smartArrow: `<path d="M12 2l8 9h-5v11H9V11H4z" fill="currentColor"/>`,
  smartTrend: stroke(`<path d="M3 17l6-6 4 4 8-8"/><path d="M15 7h6v6"/>`),
  renumber: stroke(`<circle cx="5" cy="6" r="2.5"/><circle cx="5" cy="12" r="2.5"/><circle cx="5" cy="18" r="2.5"/><path d="M10 6h11M10 12h11M10 18h11"/>`),
};

/** Sizes each icon is rendered at: ribbon sizes plus the add-in icon sizes. */
export const RENDER_SIZES = [16, 32, 64, 80];

export function svgFor(name) {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" color="${ACCENT}">${ICONS[name]}</svg>`;
}
