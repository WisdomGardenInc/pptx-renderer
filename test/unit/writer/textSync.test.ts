import { describe, expect, it } from 'vitest';
import { parseXml, SafeXmlNode } from '../../../src/parser/XmlParser';
import { parseSlide, type SlideData } from '../../../src/model/Slide';
import type { ShapeNodeData } from '../../../src/model/nodes/ShapeNode';
import { applyPlainText } from '../../../src/model/nodes/textEdit';
import { serializeSlide } from '../../../src/writer/SlideWriter';
import { A, insertOrdered, RPR_CHILD_ORDER } from '../../../src/writer/xmlEdit';

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

const BODY =
  '<a:p><a:pPr algn="ctr" marL="342900"><a:lnSpc><a:spcPct val="150000"/></a:lnSpc>' +
  '<a:buChar char="#"/></a:pPr>' +
  '<a:r><a:rPr sz="2400" b="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Hello</a:t></a:r>' +
  '<a:r><a:rPr sz="1200" i="1"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:rPr><a:t> world</a:t></a:r>' +
  '<a:endParaRPr sz="2400"/></a:p>' +
  '<a:p><a:pPr lvl="1" algn="r"><a:spcBef><a:spcPts val="6000"/></a:spcBef></a:pPr>' +
  '<a:r><a:rPr sz="1800"/><a:t>Second</a:t></a:r><a:endParaRPr sz="1800"/></a:p>';

function shape(): { slide: SlideData; node: ShapeNodeData } {
  const slide = parseSlide(
    parseXml(
      `<p:sld ${NS}><p:cSld><p:spTree>` +
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
        '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>' +
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
        `<p:txBody><a:bodyPr/><a:lstStyle/>${BODY}</p:txBody></p:sp>` +
        '</p:spTree></p:cSld></p:sld>',
    ),
    0,
    new Map(),
    'ppt/slides/slide1.xml',
  );
  return { slide, node: slide.nodes[0] as ShapeNodeData };
}

function txBody(xml: string): string {
  return xml.slice(xml.indexOf('<p:txBody>'), xml.indexOf('</p:txBody>') + 11);
}

/**
 * Every a:p in the part.
 *
 * Read from the live tree rather than the serialized string: the string is a different
 * Document, so identity assertions (a reference the model still owns) only mean anything
 * against the tree the writer actually edited.
 */
function paragraphsOf(slide: SlideData): Element[] {
  const root = slide.root!.element as Element;
  return Array.from(root.getElementsByTagName('*')).filter((el) => el.localName === 'p');
}

function child(el: Element, local: string): Element | null {
  return Array.from(el.children).find((c) => c.localName === local) ?? null;
}

describe('syncText after a plain-text edit', () => {
  it('keeps paragraph properties of every paragraph in the file', () => {
    const { slide, node } = shape();
    applyPlainText(node, 'Hello world\nSecond one');
    serializeSlide(slide);
    const paras = paragraphsOf(slide);

    expect(paras).toHaveLength(2);
    expect(child(paras[0]!, 'pPr')?.getAttribute('algn')).toBe('ctr');
    expect(child(paras[0]!, 'pPr')?.getAttribute('marL')).toBe('342900');
    expect(child(child(paras[0]!, 'pPr')!, 'buChar')).toBeTruthy();
    expect(child(paras[0]!, 'endParaRPr')?.getAttribute('sz')).toBe('2400');
    expect(child(paras[1]!, 'pPr')?.getAttribute('algn')).toBe('r');
    expect(child(paras[1]!, 'pPr')?.getAttribute('lvl')).toBe('1');
  });

  it('keeps the unedited paragraph byte-for-byte', () => {
    const { slide, node } = shape();
    const untouched = node.textBody!.paragraphs[1].properties!.element;
    applyPlainText(node, 'Hello there world\nSecond');
    serializeSlide(slide);
    const paras = paragraphsOf(slide);

    expect(child(paras[1]!, 'pPr')).toBe(untouched);
  });

  it('writes the cloned properties of a split paragraph', () => {
    const { slide, node } = shape();
    applyPlainText(node, 'Hello world\nExtra\nSecond');
    serializeSlide(slide);
    const paras = paragraphsOf(slide);

    expect(paras).toHaveLength(3);
    expect(child(paras[1]!, 'pPr')?.getAttribute('algn')).toBe('ctr');
    expect(child(child(paras[1]!, 'pPr')!, 'lnSpc')).toBeTruthy();
    expect(child(paras[1]!, 'endParaRPr')?.getAttribute('sz')).toBe('2400');
    // the copy is a distinct element, not the one still owned by paragraph one
    expect(child(paras[1]!, 'pPr')).not.toBe(child(paras[0]!, 'pPr'));
  });

  it('gives each run its own rPr even when the model shares one', () => {
    const { slide, node } = shape();
    applyPlainText(node, 'one\ntwo');
    // A host that reuses one a:rPr for every run, as a naive rewrite would.
    const shared = node.textBody!.paragraphs[0].runs[0]!.properties!;
    node.textBody!.paragraphs.forEach((p) => {
      p.runs.forEach((r) => {
        r.properties = shared;
      });
    });

    serializeSlide(slide);
    const runs = paragraphsOf(slide)
      .flatMap((p) => Array.from(p.children))
      .filter((el) => el.localName === 'r');
    expect(runs).toHaveLength(2);
    const props = runs.map((run) => child(run, 'rPr'));
    expect(props[1]).not.toBe(props[0]);
    for (const el of props) expect(el?.getAttribute('sz')).toBe('2400');
  });
});

describe('run properties keep the schema sequence', () => {
  it('places a new fill before latin instead of appending it', () => {
    const doc = parseXml(`<a:rPr xmlns:a="${A}" lang="en-US"><a:latin typeface="Arial"/></a:rPr>`);
    const rPr = doc.element!;
    const sf = rPr.ownerDocument!.createElementNS(A, 'a:solidFill');
    sf.appendChild(rPr.ownerDocument!.createElementNS(A, 'a:srgbClr'));
    insertOrdered(rPr, sf, RPR_CHILD_ORDER);
    expect(Array.from(rPr.children).map((c) => c.localName)).toEqual(['solidFill', 'latin']);
  });
});
