/**
 * `XMLSerializer` does not behave the same everywhere: browsers emit the XML
 * declaration, jsdom does not. The writer ran for a while with an unconditional
 * declaration prepended, which was invisible under jsdom and produced a doubly-declared,
 * unparseable slide part in every real browser — and an unparseable part renders as a
 * blank slide, not as an error, so nothing downstream complained either.
 *
 * These tests run against both behaviours, because the test environment only exercises
 * one of them.
 */
import { afterEach, describe, expect, it } from 'vitest';
import { parseXml } from '../../../src/parser/XmlParser';
import { parseSlide, type SlideData } from '../../../src/model/Slide';
import { serializeSlide } from '../../../src/writer/SlideWriter';

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function makeSlide(): SlideData {
  const xml =
    `<p:sld ${NS}><p:cSld><p:spTree>` +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
    '<p:sp><p:nvSpPr><p:cNvPr id="2" name="S"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
    '<p:txBody><a:bodyPr/><a:p><a:r><a:rPr/><a:t>Hi</a:t></a:r></a:p></p:txBody></p:sp>' +
    '</p:spTree></p:cSld></p:sld>';
  return parseSlide(parseXml(xml), 0, new Map(), 'ppt/slides/slide1.xml');
}

const NativeSerializer = globalThis.XMLSerializer;

afterEach(() => {
  globalThis.XMLSerializer = NativeSerializer;
});

/** Stand in for a browser, whose serializer includes the declaration. */
function useDeclaringSerializer(): void {
  globalThis.XMLSerializer = class {
    serializeToString(node: Node): string {
      return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        new NativeSerializer().serializeToString(node)
      );
    }
  } as unknown as typeof XMLSerializer;
}

function declarationCount(xml: string): number {
  return (xml.match(/<\?xml/g) ?? []).length;
}

function parses(xml: string): boolean {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  return doc.querySelector('parsererror') === null;
}

describe('serialized parts carry exactly one XML declaration', () => {
  it('adds one when the serializer omits it (jsdom-style)', () => {
    const { xml } = serializeSlide(makeSlide());
    expect(declarationCount(xml)).toBe(1);
    expect(xml.startsWith('<?xml version="1.0"')).toBe(true);
  });

  it('does not add a second when the serializer emits it (browser-style)', () => {
    useDeclaringSerializer();
    const { xml } = serializeSlide(makeSlide());
    expect(declarationCount(xml)).toBe(1);
  });

  it('stays parseable under either serializer', () => {
    expect(parses(serializeSlide(makeSlide()).xml)).toBe(true);
    useDeclaringSerializer();
    expect(parses(serializeSlide(makeSlide()).xml)).toBe(true);
  });

  it('reparses to the same shape tree under a declaring serializer', () => {
    useDeclaringSerializer();
    const { xml } = serializeSlide(makeSlide());
    const reparsed = parseSlide(parseXml(xml), 0, new Map(), 'ppt/slides/slide1.xml');
    expect(reparsed.nodes.length).toBe(1);
    expect(reparsed.nodes[0].id).toBe('2');
  });
});
