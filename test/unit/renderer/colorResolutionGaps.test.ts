/**
 * Regression cover for the three colour-resolution gaps exercised by the
 * `color-resolution-gaps` E2E case (test/e2e/testdata/cases/). That corpus is
 * git-ignored, so the behaviour is pinned here against the real
 * buildPresentation -> renderSlide pipeline rather than a mock context.
 *
 * Expected values were cross-checked against a LibreOffice render of the
 * sample deck.
 */

import { describe, expect, it } from 'vitest';
import { buildPresentation } from '../../../src/model/Presentation';
import { renderSlide } from '../../../src/renderer/SlideRenderer';
import type { PptxFiles } from '../../../src/parser/ZipParser';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

/** A filled rectangle whose solidFill is the XML under test. */
function rect(id: number, x: number, fill: string): string {
  return `
    <p:sp>
      <p:nvSpPr><p:cNvPr id="${id}" name="Rect ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="${x}" y="914400"/><a:ext cx="914400" cy="914400"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill>${fill}</a:solidFill>
      </p:spPr>
      <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>
  `;
}

function slide(shapes: string): string {
  return `
    <p:sld xmlns:a="${A_NS}" xmlns:p="${P_NS}" xmlns:r="${R_NS}">
      <p:cSld><p:spTree>
        <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr/>
        ${shapes}
      </p:spTree></p:cSld>
      <!-- PowerPoint writes this on nearly every slide; it means "inherit". -->
      <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
    </p:sld>
  `;
}

function layout(clrMapOvr: string): string {
  return `
    <p:sldLayout xmlns:a="${A_NS}" xmlns:p="${P_NS}">
      <p:cSld><p:spTree/></p:cSld>
      ${clrMapOvr}
    </p:sldLayout>
  `;
}

function rels(entries: Array<[string, string, string]>): string {
  return `<Relationships xmlns="${REL_NS}">${entries
    .map(
      ([id, type, target]) =>
        `<Relationship Id="${id}" Type="${REL_NS.replace('/package/', '/officeDocument/')}/${type}" Target="${target}"/>`,
    )
    .join('')}</Relationships>`;
}

/**
 * Deck shaped like the sample: a dark master colour map, one layout that
 * inherits it and one that overrides it back to light.
 */
function makeFiles(): PptxFiles {
  const slideIds = [1, 2, 3];

  return {
    presentation: `
      <p:presentation xmlns:p="${P_NS}" xmlns:r="${R_NS}">
        <p:sldSz cx="12192000" cy="6858000"/>
        <p:sldIdLst>
          ${slideIds.map((n) => `<p:sldId id="${255 + n}" r:id="rId${n}"/>`).join('')}
        </p:sldIdLst>
      </p:presentation>
    `,
    presentationRels: rels(slideIds.map((n) => [`rId${n}`, 'slide', `slides/slide${n}.xml`])),

    slides: new Map([
      // Layout inherits the master map: bg1 -> dk1, tx1 -> lt1.
      [
        'ppt/slides/slide1.xml',
        slide(
          rect(2, 685800, '<a:schemeClr val="bg1"/>') +
            rect(3, 1828800, '<a:schemeClr val="tx1"/>'),
        ),
      ],
      // Layout overrides the map back to light: bg1 -> lt1, tx1 -> dk1.
      [
        'ppt/slides/slide2.xml',
        slide(
          rect(2, 685800, '<a:schemeClr val="bg1"/>') +
            rect(3, 1828800, '<a:schemeClr val="tx1"/>'),
        ),
      ],
      [
        'ppt/slides/slide3.xml',
        slide(
          rect(2, 685800, '<a:prstClr val="dkBlue"/>') +
            rect(3, 1828800, '<a:prstClr val="ltGreen"/>') +
            rect(4, 2971800, '<a:prstClr val="medPurple"/>') +
            rect(5, 4114800, '<a:prstClr val="notAPresetColor"/>'),
        ),
      ],
    ]),
    slideRels: new Map([
      [
        'ppt/slides/_rels/slide1.xml.rels',
        rels([['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml']]),
      ],
      [
        'ppt/slides/_rels/slide2.xml.rels',
        rels([['rId1', 'slideLayout', '../slideLayouts/slideLayout2.xml']]),
      ],
      [
        'ppt/slides/_rels/slide3.xml.rels',
        rels([['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml']]),
      ],
    ]),

    slideLayouts: new Map([
      [
        'ppt/slideLayouts/slideLayout1.xml',
        layout('<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'),
      ],
      [
        'ppt/slideLayouts/slideLayout2.xml',
        layout(
          '<p:clrMapOvr><a:overrideClrMapping bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"/></p:clrMapOvr>',
        ),
      ],
    ]),
    slideLayoutRels: new Map([
      [
        'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
        rels([['rId1', 'slideMaster', '../slideMasters/slideMaster1.xml']]),
      ],
      [
        'ppt/slideLayouts/_rels/slideLayout2.xml.rels',
        rels([['rId1', 'slideMaster', '../slideMasters/slideMaster1.xml']]),
      ],
    ]),

    slideMasters: new Map([
      [
        'ppt/slideMasters/slideMaster1.xml',
        `
          <p:sldMaster xmlns:a="${A_NS}" xmlns:p="${P_NS}">
            <p:cSld><p:spTree/></p:cSld>
            <p:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1"
                      accent2="accent2" accent3="accent3" accent4="accent4"
                      accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
          </p:sldMaster>
        `,
      ],
    ]),
    slideMasterRels: new Map([
      [
        'ppt/slideMasters/_rels/slideMaster1.xml.rels',
        rels([['rId1', 'theme', '../theme/theme1.xml']]),
      ],
    ]),

    themes: new Map([
      [
        'ppt/theme/theme1.xml',
        `
          <a:theme xmlns:a="${A_NS}">
            <a:themeElements>
              <a:clrScheme name="Office">
                <a:dk1><a:srgbClr val="0A1A2F"/></a:dk1>
                <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                <a:dk2><a:srgbClr val="14263D"/></a:dk2>
                <a:lt2><a:srgbClr val="F0F2F5"/></a:lt2>
                <a:accent1><a:srgbClr val="1E81FD"/></a:accent1>
              </a:clrScheme>
              <a:fontScheme name="Office">
                <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
                <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
              </a:fontScheme>
              <a:fmtScheme name="Office">
                <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
                <a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
                <a:effectStyleLst/>
              </a:fmtScheme>
            </a:themeElements>
          </a:theme>
        `,
      ],
    ]),

    media: new Map(),
    charts: new Map(),
    diagramDrawings: new Map(),
  };
}

/** Hex fills of the rendered shape paths, in shape order. */
function renderedFills(slideIndex: number): string[] {
  const presentation = buildPresentation(makeFiles());
  const handle = renderSlide(presentation, presentation.slides[slideIndex]);
  const fills: string[] = [];
  handle.element.querySelectorAll('path').forEach((path) => {
    const fill = path.getAttribute('fill');
    if (fill && fill !== 'none') fills.push(fill.toUpperCase());
  });
  handle.dispose();
  return fills;
}

describe('colour resolution gaps', () => {
  it('maps scheme colours through the master clrMap when no layout overrides it', () => {
    expect(renderedFills(0)).toEqual(['#0A1A2F', '#FFFFFF']);
  });

  it('lets a layout overrideClrMapping reach a slide that declares masterClrMapping', () => {
    // The slide's own masterClrMapping means "inherit", so the layout override
    // wins and the mapping flips: bg1 -> lt1, tx1 -> dk1. Reading it as "jump to
    // the master" would render this slide identically to the one above.
    expect(renderedFills(1)).toEqual(['#FFFFFF', '#0A1A2F']);
  });

  it('resolves the abbreviated prstClr names and greys out unknown ones', () => {
    // ECMA-376 ST_PresetColorVal spells these dk/lt/med, not dark/light/medium.
    expect(renderedFills(2)).toEqual(['#00008B', '#90EE90', '#9370DB', '#808080']);
  });
});
