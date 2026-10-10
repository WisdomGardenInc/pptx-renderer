import { describe, expect, it } from 'vitest';
import { parseXml } from '../../../../src/parser/XmlParser';
import { parseShapeNode } from '../../../../src/model/nodes/ShapeNode';
import type { ShapeNodeData } from '../../../../src/model/nodes/ShapeNode';
import {
  applyPlainText,
  paragraphPlainText,
  readPlainText,
} from '../../../../src/model/nodes/textEdit';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

function box(txBody: string): ShapeNodeData {
  return parseShapeNode(
    parseXml(
      `<p:sp xmlns:a="${A}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">` +
        '<p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
        '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
        `<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>${txBody}</p:txBody>` +
        '</p:sp>',
    ),
  );
}

const MIXED =
  '<a:p><a:pPr algn="ctr" marL="342900"><a:lnSpc><a:spcPct val="150000"/></a:lnSpc>' +
  '<a:buChar char="#"/></a:pPr>' +
  '<a:r><a:rPr sz="2400" b="1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>Hello</a:t></a:r>' +
  '<a:r><a:rPr sz="1200" i="1"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:rPr><a:t> world</a:t></a:r>' +
  '<a:endParaRPr sz="2400"/></a:p>' +
  '<a:p><a:pPr lvl="1" algn="r"><a:spcBef><a:spcPts val="6000"/></a:spcBef></a:pPr>' +
  '<a:r><a:rPr sz="1800"/><a:t>Second</a:t></a:r></a:p>';

describe('applyPlainText', () => {
  it('keeps paragraph properties of an edited paragraph', () => {
    const shape = box(MIXED);
    applyPlainText(shape, 'Hello there world');

    const [first] = shape.textBody!.paragraphs;
    expect(first.properties?.attr('algn')).toBe('ctr');
    expect(first.properties?.numAttr('marL')).toBe(342900);
    expect(first.properties?.child('buChar').attr('char')).toBe('#');
    expect(first.properties?.child('lnSpc').child('spcPct').attr('val')).toBe('150000');
    expect(first.endParaRPr?.numAttr('sz')).toBe(2400);
  });

  it('keeps the level and spacing of the paragraph that was not edited', () => {
    const shape = box(MIXED);
    applyPlainText(shape, 'Hello world\nSecond one');

    const [, second] = shape.textBody!.paragraphs;
    expect(second.level).toBe(1);
    expect(second.properties?.attr('algn')).toBe('r');
    expect(second.properties?.child('spcBef').child('spcPts').attr('val')).toBe('6000');
    expect(paragraphPlainText(second)).toBe('Second one');
  });

  it('replaces only the edited span of a mixed-style paragraph', () => {
    const shape = box(MIXED);
    applyPlainText(shape, 'Hello brave world');

    const runs = shape.textBody!.paragraphs[0].runs;
    expect(runs.map((r) => r.text)).toEqual(['Hello', ' brave world']);
    expect(paragraphPlainText(shape.textBody!.paragraphs[0])).toBe('Hello brave world');
    // the run the edit never reached keeps italic blue
    expect(runs[1]!.properties?.numAttr('sz')).toBe(1200);
    expect(runs[1]!.properties?.attr('i')).toBe('1');
    expect(runs[1]!.properties?.child('solidFill').child('srgbClr').attr('val')).toBe('0000FF');
    // and the one before it keeps bold red
    expect(runs[0]!.properties?.numAttr('sz')).toBe(2400);
    expect(runs[0]!.properties?.attr('b')).toBe('1');
  });

  it('leaves an unchanged paragraph object alone, soft breaks included', () => {
    const shape = box(
      '<a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="2000"/><a:t>one</a:t></a:r>' +
        '<a:br><a:rPr sz="2000"/></a:br><a:r><a:rPr sz="900"/><a:t>two</a:t></a:r></a:p>' +
        '<a:p><a:pPr algn="l"/><a:r><a:t>next</a:t></a:r></a:p>',
    );
    const before = shape.textBody!.paragraphs[0];
    expect(readPlainText(shape)).toBe('one\ntwo\nnext');

    applyPlainText(shape, 'one\ntwo\nnext edited');

    expect(shape.textBody!.paragraphs[0]).toBe(before);
    expect(before.runs.map((r) => r.text)).toEqual(['one', '\n', 'two']);
  });

  it('splits a paragraph so the new line inherits its styling', () => {
    const shape = box(
      '<a:p><a:pPr lvl="2" algn="just"><a:spcAft><a:spcPts val="4000"/></a:spcAft></a:pPr>' +
        '<a:r><a:rPr sz="1000" u="sng"/><a:t>Alpha</a:t></a:r>' +
        '<a:endParaRPr sz="1000"/></a:p>',
    );
    applyPlainText(shape, 'Alpha\nBeta');

    const [first, second] = shape.textBody!.paragraphs;
    expect(paragraphPlainText(first)).toBe('Alpha');
    expect(paragraphPlainText(second)).toBe('Beta');
    // the original keeps its element references, the copy gets its own
    expect(first.properties?.element).not.toBe(second.properties?.element);
    expect(second.properties?.attr('algn')).toBe('just');
    expect(second.properties?.child('spcAft').child('spcPts').attr('val')).toBe('4000');
    expect(second.level).toBe(2);
    expect(second.endParaRPr?.numAttr('sz')).toBe(1000);
    expect(second.runs[0]!.properties?.attr('u')).toBe('sng');
  });

  it('never points two runs at one a:rPr element', () => {
    const shape = box(
      '<a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="1600"/><a:t>Alpha</a:t></a:r></a:p>',
    );
    applyPlainText(shape, 'Alpha\nAlpha');

    const [first, second] = shape.textBody!.paragraphs;
    const els = [first, second].map((p) => p.properties?.element);
    expect(els[0]).not.toBe(els[1]);
    const runEls = [first.runs[0]!.properties?.element, second.runs[0]!.properties?.element];
    expect(runEls[0]).not.toBe(runEls[1]);
    expect(runEls[1]!.isEqualNode(runEls[0]!)).toBe(true);
  });

  it('keeps one paragraph with its properties when everything is deleted', () => {
    const shape = box(MIXED);
    applyPlainText(shape, '');

    expect(shape.textBody!.paragraphs).toHaveLength(1);
    expect(shape.textBody!.paragraphs[0].properties?.attr('algn')).toBe('ctr');
    expect(shape.textBody!.paragraphs[0].runs).toEqual([]);
  });

  it('adds paragraphs when the body was empty', () => {
    const shape = box('');
    applyPlainText(shape, 'one\ntwo');
    expect(readPlainText(shape)).toBe('one\ntwo');
    expect(shape.textBody!.paragraphs[1].runs[0]!.text).toBe('two');
  });
});
