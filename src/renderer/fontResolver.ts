/**
 * Font resolution helpers shared by text-like renderers.
 */

import type { RenderContext } from './RenderContext';

type ThemeFontSlot = 'lt' | 'ea' | 'cs';
type ThemeFontKey = 'latin' | 'ea' | 'cs';
type LanguageHint = string | null | undefined;

const THEME_FONT_REF = /^\+(mj|mn)-(lt|ea|cs)$/;
const THEME_FONT_SLOT_MAP: Record<ThemeFontSlot, ThemeFontKey> = {
  lt: 'latin',
  ea: 'ea',
  cs: 'cs',
};

const EAST_ASIAN_SCRIPT_FALLBACKS = ['Hans', 'Hant', 'Jpan', 'Hang'];

function normalizeLanguageHints(hints?: LanguageHint | LanguageHint[]): string[] {
  if (Array.isArray(hints)) {
    return hints.filter((hint): hint is string => typeof hint === 'string' && hint.length > 0);
  }
  return typeof hints === 'string' && hints.length > 0 ? [hints] : [];
}

function scriptFromLanguage(lang: string): string | undefined {
  const normalized = lang.toLowerCase();
  if (normalized.startsWith('zh')) {
    return /-(tw|hk|mo)\b/.test(normalized) ? 'Hant' : 'Hans';
  }
  if (normalized.startsWith('ja')) return 'Jpan';
  if (normalized.startsWith('ko')) return 'Hang';
  if (normalized.startsWith('ar')) return 'Arab';
  if (normalized.startsWith('he')) return 'Hebr';
  if (normalized.startsWith('th')) return 'Thai';
  if (normalized.startsWith('hi') || normalized.startsWith('mr') || normalized.startsWith('ne')) {
    return 'Deva';
  }
  return undefined;
}

function resolveScriptFont(
  scripts: Record<string, string> | undefined,
  hints?: LanguageHint | LanguageHint[],
): string | undefined {
  if (!scripts) return undefined;
  for (const hint of normalizeLanguageHints(hints)) {
    const script = scriptFromLanguage(hint);
    if (script && scripts[script]) return scripts[script];
  }
  for (const script of EAST_ASIAN_SCRIPT_FALLBACKS) {
    if (scripts[script]) return scripts[script];
  }
  return undefined;
}

function resolveThemeFontName(
  typeface: string,
  ctx: RenderContext,
  languageHints?: LanguageHint | LanguageHint[],
): string {
  const match = typeface.match(THEME_FONT_REF);
  if (!match) return typeface;

  const scheme = match[1];
  const slot = match[2] as ThemeFontSlot;
  const fonts = scheme === 'mj' ? ctx.theme.majorFont : ctx.theme.minorFont;
  const key = THEME_FONT_SLOT_MAP[slot];
  const direct = fonts[key];
  if (direct) return direct;
  if (slot === 'ea') {
    const scriptFont = resolveScriptFont(fonts.scripts, languageHints);
    if (scriptFont) return scriptFont;
  }
  return fonts.latin || fonts.ea || fonts.cs || typeface;
}

/** Embedded faces are only consulted when the host opted in; see `embeddedFontsEnabled`. */
function embeddedFamilyFor(typeface: string, ctx: RenderContext): string | undefined {
  if (!ctx.embeddedFontsEnabled) return undefined;
  return ctx.presentation.embeddedFontFamilies?.get(typeface.trim().toLowerCase());
}

/**
 * Resolve theme font placeholder references like "+mj-lt" or "+mn-ea".
 */
export function resolveThemeFont(
  typeface: string,
  ctx: RenderContext,
  languageHints?: LanguageHint | LanguageHint[],
): string {
  const resolved = resolveThemeFontName(typeface, ctx, languageHints);
  const embeddedFamily = embeddedFamilyFor(resolved, ctx);
  if (!embeddedFamily) return resolved;
  ctx.usedEmbeddedFontFamilies?.add(embeddedFamily);
  return embeddedFamily;
}

export function resolveThemeFontStack(
  typefaces: (string | undefined)[],
  ctx: RenderContext,
  languageHints?: LanguageHint | LanguageHint[],
): string[] {
  const seen = new Set<string>();
  const stack: string[] = [];
  for (const typeface of typefaces) {
    if (!typeface) continue;
    const resolved = resolveThemeFontName(typeface, ctx, languageHints).trim();
    const embedded = embeddedFamilyFor(resolved, ctx);
    if (embedded) ctx.usedEmbeddedFontFamilies?.add(embedded);
    for (const font of embedded ? [embedded, resolved] : [resolved]) {
      const key = font.toLowerCase();
      if (!font || seen.has(key)) continue;
      seen.add(key);
      stack.push(font);
    }
  }
  return stack;
}

const CSS_GENERIC_FONT_FAMILIES = new Set([
  'serif',
  'sans-serif',
  'monospace',
  'cursive',
  'fantasy',
  'system-ui',
  'ui-serif',
  'ui-sans-serif',
  'ui-monospace',
  'emoji',
  'math',
  'fangsong',
]);

/**
 * Which system faces to fall back to for a CJK family, keyed by the style its name
 * implies. Decks routinely name fonts the viewer does not have installed, and a bare
 * `font-family: 汉仪小隶书简` leaves the browser on its default standard font — a modern
 * sans at best, tofu on a system with no CJK coverage.
 */
type CjkFontStyle = 'sans' | 'serif' | 'kai';

const CJK_FALLBACKS: Record<CjkFontStyle, string[]> = {
  sans: [
    'PingFang SC',
    'Hiragino Sans GB',
    'Microsoft YaHei',
    'Noto Sans CJK SC',
    'Source Han Sans SC',
    'Arial Unicode MS',
    'sans-serif',
  ],
  serif: [
    'Songti SC',
    'SimSun',
    'Noto Serif CJK SC',
    'Source Han Serif SC',
    'Arial Unicode MS',
    'serif',
  ],
  kai: ['Kaiti SC', 'KaiTi', 'STKaiti', 'Songti SC', 'SimSun', 'Noto Serif CJK SC', 'serif'],
};

/** Brush-script families (楷体, 行书, …) — the Kai faces are the closest system match. */
const CJK_KAI_NAME = /楷|行[书書]|草[书書]/;

/** Printed serif families (宋体, 明朝, 隶书, 篆书, …). */
const CJK_SERIF_NAME = /宋|明[体體朝]|隶|篆|魏碑/;

/** Han, kana and hangul — a family whose own name uses them is a CJK face. */
const CJK_NAME_CHARS = /[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\uac00-\ud7af]/;

/**
 * Romanised CJK style words, for families such as "A-OTF Kaisho MCBK1 Pro MCBK1"
 * that contain no CJK character. Only words that never appear in a Latin family
 * name belong here — "gothic" and "song" are far too common to match on.
 */
const CJK_LATIN_NAME_HINTS: Array<[RegExp, CjkFontStyle]> = [
  [/\bkaisho\b/i, 'kai'],
  [/\b(?:mincho|myeongjo)\b/i, 'serif'],
  [/\bmaru\s?gothic\b/i, 'sans'],
];

/**
 * CJK faces spelled in Latin, which `CJK_NAME_CHARS` cannot recognise. Their names
 * carry no style keyword either ("SimSun" is a 宋体), so each one states its style.
 */
const CJK_LATIN_FAMILIES = new Map<string, CjkFontStyle>([
  ['microsoft yahei', 'sans'],
  ['microsoft yahei ui', 'sans'],
  ['msyh', 'sans'],
  ['dengxian', 'sans'],
  ['simhei', 'sans'],
  ['stheiti', 'sans'],
  ['stxihei', 'sans'],
  ['heiti sc', 'sans'],
  ['heiti tc', 'sans'],
  ['pingfang sc', 'sans'],
  ['pingfang tc', 'sans'],
  ['pingfang hk', 'sans'],
  ['hiragino sans gb', 'sans'],
  ['noto sans cjk sc', 'sans'],
  ['source han sans sc', 'sans'],
  ['ms gothic', 'sans'],
  ['ms pgothic', 'sans'],
  ['meiryo', 'sans'],
  ['yu gothic', 'sans'],
  ['malgun gothic', 'sans'],
  ['gulim', 'sans'],
  ['dotum', 'sans'],
  ['simsun', 'serif'],
  ['nsimsun', 'serif'],
  ['songti sc', 'serif'],
  ['songti tc', 'serif'],
  ['stsong', 'serif'],
  ['fangsong', 'serif'],
  ['stfangsong', 'serif'],
  ['noto serif cjk sc', 'serif'],
  ['source han serif sc', 'serif'],
  ['mingliu', 'serif'],
  ['pmingliu', 'serif'],
  ['ms mincho', 'serif'],
  ['ms pmincho', 'serif'],
  ['yu mincho', 'serif'],
  ['batang', 'serif'],
  ['simkai', 'kai'],
  ['kaiti', 'kai'],
  ['kaiti sc', 'kai'],
  ['kaiti tc', 'kai'],
  ['stkaiti', 'kai'],
]);

const FONT_FAMILY_ALIASES: Record<string, string[]> = {
  calibri: ['Calibri', 'Aptos', 'Carlito', 'system-ui', 'Arial', 'Helvetica', 'sans-serif'],
  'calibri light': [
    'Calibri Light',
    'Aptos Display',
    'Aptos',
    'Carlito',
    'system-ui',
    'Arial',
    'Helvetica',
    'sans-serif',
  ],
  aptos: ['Aptos', 'system-ui', 'Arial', 'Helvetica', 'sans-serif'],
  'aptos display': ['Aptos Display', 'Aptos', 'system-ui', 'Arial', 'Helvetica', 'sans-serif'],
  'microsoft yahei': ['Microsoft YaHei', '微软雅黑'],
  'microsoft yahei ui': ['Microsoft YaHei UI', 'Microsoft YaHei', '微软雅黑'],
  微软雅黑: ['微软雅黑', 'Microsoft YaHei'],
  dengxian: ['DengXian', '等线'],
  等线: ['等线', 'DengXian'],
  simhei: ['SimHei', '黑体'],
  黑体: ['黑体', 'SimHei'],
  'heiti sc': ['Heiti SC', '黑体', 'SimHei'],
};

function normalizeFontFamilyName(fontFamily: string): string {
  return fontFamily
    .trim()
    .replace(/^['"]|['"]$/g, '')
    .toLowerCase();
}

function cssFontFamilyToken(fontFamily: string): string {
  const normalized = normalizeFontFamilyName(fontFamily);
  if (CSS_GENERIC_FONT_FAMILIES.has(normalized)) {
    return normalized;
  }
  return `"${fontFamily.trim().replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
}

function expandFontFamilyAliases(fontFamily: string): string[] {
  const normalized = normalizeFontFamilyName(fontFamily);
  return FONT_FAMILY_ALIASES[normalized] ?? [fontFamily.trim()];
}

/** True for families that render Chinese, Japanese or Korean text. */
function isCjkFontFamily(fontFamily: string): boolean {
  const normalized = normalizeFontFamilyName(fontFamily);
  if (CJK_NAME_CHARS.test(normalized) || CJK_LATIN_FAMILIES.has(normalized)) return true;
  return CJK_LATIN_NAME_HINTS.some(([pattern]) => pattern.test(normalized));
}

/** Pick the fallback chain that best matches the style a CJK family name implies. */
function cjkFontStyle(fontFamily: string): CjkFontStyle {
  const normalized = normalizeFontFamilyName(fontFamily);
  const known = CJK_LATIN_FAMILIES.get(normalized);
  if (known) return known;
  if (CJK_KAI_NAME.test(fontFamily)) return 'kai';
  if (CJK_SERIF_NAME.test(fontFamily)) return 'serif';
  const hinted = CJK_LATIN_NAME_HINTS.find(([pattern]) => pattern.test(normalized));
  return hinted ? hinted[1] : 'sans';
}

export function cssFontFamilyStack(fontFamily: string | string[]): string {
  const baseFonts = Array.isArray(fontFamily)
    ? fontFamily.flatMap(expandFontFamilyAliases)
    : expandFontFamilyAliases(fontFamily);
  // Any CJK family in the stack earns a fallback chain — the requested face is often
  // a foundry font the viewer does not have, and without this the browser is left on
  // its default standard font.
  const cjkFamily = baseFonts.find(isCjkFontFamily);
  const stack = cjkFamily ? [...baseFonts, ...CJK_FALLBACKS[cjkFontStyle(cjkFamily)]] : baseFonts;
  const seen = new Set<string>();
  const unique = stack.filter((font) => {
    const key = normalizeFontFamilyName(font);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
  return unique.map(cssFontFamilyToken).join(', ');
}
