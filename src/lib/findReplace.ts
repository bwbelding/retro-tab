// Find and update across the deck: finding matches in every text on every slide (pure, unit
// tested). Reading the deck and replacing are in findActions.ts.

/** One piece of text in the deck: a shape's text or a table cell's. */
export interface TextSource {
  slideId: string;
  slideIndex: number;
  /** Shape ids from the slide's top level down to the shape. */
  path: string[];
  /** For table cells. */
  cell?: { row: number; column: number };
  /** e.g. Text box “Revenue”, or Table, row 2. */
  where: string;
  text: string;
}

export interface Match {
  /** Stable within one search, for remembering which matches you unticked. */
  key: string;
  source: TextSource;
  start: number;
  length: number;
  /** Text around the match, for showing it in context. */
  before: string;
  found: string;
  after: string;
}

export interface FindOptions {
  matchCase: boolean;
  /** Only where the text isn't part of a longer word or number ("4.2" matches "$4.2M", not "14.25"). */
  wholeWord: boolean;
}

const WORD = /[\p{L}\p{N}]/u;
const DIGIT = /\p{N}/u;

/**
 * Whether a match's edge stands apart from the text beside it (`dir` -1 looks left, 1 right). A
 * word can't touch another letter or digit. A number can't touch another digit, or a decimal
 * point or comma with digits after it, but can touch units: "4.2" matches "$4.2M", not "14.25".
 */
function apart(text: string, edge: number, dir: -1 | 1, edgeChar: string): boolean {
  const next = text[edge + dir] ?? "";
  if (DIGIT.test(edgeChar)) return !DIGIT.test(next) && !(/[.,]/.test(next) && DIGIT.test(text[edge + 2 * dir] ?? ""));
  if (WORD.test(edgeChar)) return !WORD.test(next);
  return true;
}
const CONTEXT = 24;

/** Every match of `query` in the sources, in deck order. */
export function findMatches(sources: TextSource[], query: string, options: FindOptions): Match[] {
  if (!query) return [];
  const needle = options.matchCase ? query : query.toLowerCase();
  const matches: Match[] = [];
  for (const source of sources) {
    const hay = options.matchCase ? source.text : source.text.toLowerCase();
    for (let at = hay.indexOf(needle); at >= 0; at = hay.indexOf(needle, at + needle.length)) {
      const end = at + needle.length;
      if (options.wholeWord && !(apart(source.text, at, -1, needle[0]) && apart(source.text, end - 1, 1, needle.at(-1)!))) continue;
      const before = source.text.slice(Math.max(0, at - CONTEXT), at).replace(/\s+/g, " ");
      const after = source.text.slice(end, end + CONTEXT).replace(/\s+/g, " ");
      matches.push({
        key: `${source.slideId}/${source.path.join("/")}/${source.cell ? `${source.cell.row},${source.cell.column}` : ""}@${at}`,
        source,
        start: at,
        length: needle.length,
        before: (at > CONTEXT ? "…" : "") + before,
        found: source.text.slice(at, end),
        after: after + (end + CONTEXT < source.text.length ? "…" : ""),
      });
    }
  }
  return matches;
}

/** A text with some of its matches replaced (for table cells, which are rewritten whole). */
export function replaceIn(text: string, matches: Pick<Match, "start" | "length">[], replacement: string): string {
  let out = text;
  for (const m of [...matches].sort((a, b) => b.start - a.start)) out = out.slice(0, m.start) + replacement + out.slice(m.start + m.length);
  return out;
}
