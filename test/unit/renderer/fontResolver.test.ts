import { describe, expect, it } from 'vitest';
import {
  cssFontFamilyStack,
  resolveThemeFont,
  resolveThemeFontStack,
} from '../../../src/renderer/fontResolver';
import { createMockRenderContext } from '../helpers/mockContext';

describe('fontResolver', () => {
  it('returns literal typefaces that are not theme placeholders', () => {
    const ctx = createMockRenderContext();

    expect(resolveThemeFont('Arial', ctx)).toBe('Arial');
  });

  it('resolves theme placeholders through direct major and minor slots', () => {
    const ctx = createMockRenderContext();
    ctx.theme.majorFont = { latin: 'Major Latin', ea: 'Major EA', cs: 'Major CS' };
    ctx.theme.minorFont = { latin: 'Minor Latin', ea: 'Minor EA', cs: 'Minor CS' };

    expect(resolveThemeFont('+mj-lt', ctx)).toBe('Major Latin');
    expect(resolveThemeFont('+mj-ea', ctx)).toBe('Major EA');
    expect(resolveThemeFont('+mj-cs', ctx)).toBe('Major CS');
    expect(resolveThemeFont('+mn-lt', ctx)).toBe('Minor Latin');
    expect(resolveThemeFont('+mn-ea', ctx)).toBe('Minor EA');
    expect(resolveThemeFont('+mn-cs', ctx)).toBe('Minor CS');
  });

  it('uses language-specific script fonts for East Asian theme placeholders', () => {
    const ctx = createMockRenderContext();
    ctx.theme.majorFont = {
      latin: 'Major Latin',
      ea: '',
      cs: '',
      scripts: {
        Hans: 'Hans Font',
        Hant: 'Hant Font',
        Jpan: 'Jpan Font',
        Hang: 'Hang Font',
        Arab: 'Arab Font',
        Hebr: 'Hebr Font',
        Thai: 'Thai Font',
        Deva: 'Deva Font',
      },
    };

    expect(resolveThemeFont('+mj-ea', ctx, 'zh-CN')).toBe('Hans Font');
    expect(resolveThemeFont('+mj-ea', ctx, 'zh-TW')).toBe('Hant Font');
    expect(resolveThemeFont('+mj-ea', ctx, 'ja-JP')).toBe('Jpan Font');
    expect(resolveThemeFont('+mj-ea', ctx, 'ko-KR')).toBe('Hang Font');
    expect(resolveThemeFont('+mj-ea', ctx, 'ar-SA')).toBe('Arab Font');
    expect(resolveThemeFont('+mj-ea', ctx, 'he-IL')).toBe('Hebr Font');
    expect(resolveThemeFont('+mj-ea', ctx, 'th-TH')).toBe('Thai Font');
    expect(resolveThemeFont('+mj-ea', ctx, 'hi-IN')).toBe('Deva Font');
  });

  it('falls back through script table, latin, ea, cs, and finally the placeholder text', () => {
    const ctx = createMockRenderContext();
    ctx.theme.majorFont = {
      latin: '',
      ea: '',
      cs: '',
      scripts: { Jpan: 'Jpan Fallback' },
    };
    ctx.theme.minorFont = { latin: '', ea: 'Minor EA Fallback', cs: 'Minor CS Fallback' };

    expect(resolveThemeFont('+mj-ea', ctx, 'en-US')).toBe('Jpan Fallback');
    expect(resolveThemeFont('+mn-lt', ctx)).toBe('Minor EA Fallback');

    ctx.theme.minorFont = { latin: '', ea: '', cs: 'Minor CS Fallback' };
    expect(resolveThemeFont('+mn-lt', ctx)).toBe('Minor CS Fallback');

    ctx.theme.minorFont = { latin: '', ea: '', cs: '' };
    expect(resolveThemeFont('+mn-lt', ctx)).toBe('+mn-lt');
  });

  it('builds a resolved font stack while filtering empty and duplicate typefaces', () => {
    const ctx = createMockRenderContext();
    ctx.theme.minorFont = { latin: 'Calibri', ea: 'Microsoft YaHei', cs: '' };

    expect(resolveThemeFontStack(['+mn-lt', '', undefined, 'Calibri', '+mn-ea'], ctx)).toEqual([
      'Calibri',
      'Microsoft YaHei',
    ]);
  });

  it('keeps the original family as fallback when an embedded face cannot load', () => {
    const ctx = createMockRenderContext();
    ctx.presentation.embeddedFontFamilies = new Map([['example sans', '__pptx_embedded_1_0']]);
    ctx.usedEmbeddedFontFamilies = new Set();
    ctx.embeddedFontsEnabled = true;

    expect(resolveThemeFontStack(['Example Sans'], ctx)).toEqual([
      '__pptx_embedded_1_0',
      'Example Sans',
    ]);
    expect(ctx.usedEmbeddedFontFamilies).toEqual(new Set(['__pptx_embedded_1_0']));
  });

  it('ignores embedded faces unless the host opts in', () => {
    // PowerPoint subsets embedded faces to the glyphs it believes are used. A subset that misses a
    // glyph makes the browser fall back per character, so one run renders in two typefaces at two
    // apparent sizes. Staying on host fonts keeps a run internally consistent.
    const ctx = createMockRenderContext();
    ctx.presentation.embeddedFontFamilies = new Map([['example sans', '__pptx_embedded_1_0']]);
    ctx.usedEmbeddedFontFamilies = new Set();

    expect(resolveThemeFontStack(['Example Sans'], ctx)).toEqual(['Example Sans']);
    expect(resolveThemeFont('Example Sans', ctx)).toBe('Example Sans');
    expect(ctx.usedEmbeddedFontFamilies).toEqual(new Set());
  });

  it('serializes CSS font family stacks with aliases, CJK fallbacks, generics, and escaping', () => {
    expect(cssFontFamilyStack('Calibri')).toBe(
      '"Calibri", "Aptos", "Carlito", system-ui, "Arial", "Helvetica", sans-serif',
    );
    expect(cssFontFamilyStack(['Calibri', 'sans-serif'])).toBe(
      '"Calibri", "Aptos", "Carlito", system-ui, "Arial", "Helvetica", sans-serif',
    );
    expect(cssFontFamilyStack('Calibri Light')).toBe(
      '"Calibri Light", "Aptos Display", "Aptos", "Carlito", system-ui, "Arial", "Helvetica", sans-serif',
    );
    expect(cssFontFamilyStack('Aptos')).toBe(
      '"Aptos", system-ui, "Arial", "Helvetica", sans-serif',
    );
    expect(cssFontFamilyStack('微软雅黑')).toContain('"PingFang SC"');
    expect(cssFontFamilyStack('A "Quoted" \\ Font')).toBe('"A \\"Quoted\\" \\\\ Font"');
  });

  describe('CJK fallback chains', () => {
    it('appends a fallback chain to any CJK-named family, not just known aliases', () => {
      // 汉仪小隶书简 is a foundry font almost nobody has installed. Without a
      // fallback the browser is left on its default standard font.
      const stack = cssFontFamilyStack('汉仪小隶书简');
      expect(stack.startsWith('"汉仪小隶书简", ')).toBe(true);
      expect(stack).toContain('"Songti SC"');
      expect(stack.endsWith('serif')).toBe(true);
    });

    it('picks the serif chain for 宋/明/隶 style names', () => {
      for (const family of ['宋体', '华文中宋', '汉仪小隶书简', '方正小篆体']) {
        const stack = cssFontFamilyStack(family);
        expect(stack).toContain('"Songti SC"');
        expect(stack).not.toContain('"PingFang SC"');
      }
    });

    it('picks the Kai chain for brush-script names', () => {
      for (const family of ['楷体', '华文行楷', 'STKaiti', 'SimKai']) {
        expect(cssFontFamilyStack(family)).toContain('"Kaiti SC"');
      }
    });

    it('picks the sans chain for hei/yahei style names', () => {
      for (const family of ['微软雅黑', '黑体', '方正兰亭黑', '华文细黑']) {
        const stack = cssFontFamilyStack(family);
        expect(stack).toContain('"PingFang SC"');
        expect(stack.endsWith('sans-serif')).toBe(true);
      }
    });

    it('classifies Latin-spelled CJK families by an explicit style table', () => {
      expect(cssFontFamilyStack('SimSun')).toContain('"Songti SC"');
      expect(cssFontFamilyStack('MS Mincho')).toContain('"Songti SC"');
      expect(cssFontFamilyStack('Batang')).toContain('"Songti SC"');
      expect(cssFontFamilyStack('Meiryo')).toContain('"PingFang SC"');
      expect(cssFontFamilyStack('Malgun Gothic')).toContain('"PingFang SC"');
    });

    it('keeps the requested families first and drops duplicates', () => {
      const stack = cssFontFamilyStack(['Arial', '宋体']);
      expect(stack.startsWith('"Arial", "宋体", ')).toBe(true);
      expect(stack.match(/"SimSun"/g)).toHaveLength(1);
      expect(cssFontFamilyStack('SimSun').match(/"SimSun"/g)).toHaveLength(1);
    });

    it('recognises romanised CJK style words in otherwise Latin family names', () => {
      expect(cssFontFamilyStack('A-OTF Kaisho MCBK1 Pro MCBK1')).toContain('"Kaiti SC"');
      expect(cssFontFamilyStack('A-OTF Ryumin Pr6N Mincho')).toContain('"Songti SC"');
      // "Gothic" and "Song" are ordinary Latin family words and must not match.
      expect(cssFontFamilyStack('Century Gothic')).toBe('"Century Gothic"');
      expect(cssFontFamilyStack('Song Sans')).toBe('"Song Sans"');
    });

    it('leaves non-CJK families untouched', () => {
      expect(cssFontFamilyStack('Arial')).toBe('"Arial"');
      expect(cssFontFamilyStack('Times New Roman')).toBe('"Times New Roman"');
      expect(cssFontFamilyStack('Aptos')).not.toContain('PingFang');
    });
  });
});
