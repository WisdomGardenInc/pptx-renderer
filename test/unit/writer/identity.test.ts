/**
 * `cNvPr/@id` identity across a save.
 *
 * The editor builds a new shape by parsing a seed XML string and stamping the typed
 * `id` onto the resulting node — the markup keeps the seed's own id. Nothing used to
 * carry that typed id back, so every added shape was written out under the same id;
 * on reload they collided in the host's `Map<id, …>` and all but the last became
 * impossible to select. The file still opened, and the shapes still drew, which is why
 * it looked like "the ones I added last time can't be edited any more".
 */
import { describe, expect, it } from 'vitest';
import { parseXml } from '../../../src/parser/XmlParser';
import { parseSlide, type SlideData } from '../../../src/model/Slide';
import { parseShapeNode, type ShapeNodeData } from '../../../src/model/nodes/ShapeNode';
import { serializeSlide } from '../../../src/writer/SlideWriter';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS =
  `xmlns:a="${A}" ` +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function existingShape(id: string, name: string): string {
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr></p:sp>'
  );
}

function slideWith(body: string): SlideData {
  const xml =
    `<p:sld ${NS}><p:cSld><p:spTree>` +
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
    body +
    '</p:spTree></p:cSld></p:sld>';
  return parseSlide(parseXml(xml), 0, new Map(), 'ppt/slides/slide1.xml');
}

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';

/** A node built the way the editor builds one: seed markup, typed id stamped on top. */
function addedShape(typedId: string): ShapeNodeData {
  const node = parseShapeNode(
    parseXml(
      `<p:sp xmlns:a="${A}" xmlns:p="${P}">` +
        '<p:nvSpPr><p:cNvPr id="0" name="Rectangle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
        '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm></p:spPr></p:sp>',
    ),
  );
  node.id = typedId;
  node.name = 'Rectangle';
  return node;
}

function idsIn(xml: string): string[] {
  return [...xml.matchAll(/<p:cNvPr id="([^"]+)"/g)].map((m) => m[1]);
}

describe('cNvPr ids survive a save', () => {
  it('assigns distinct ids to several added shapes', () => {
    const slide = slideWith(existingShape('2', 'A') + existingShape('3', 'B'));
    slide.nodes.push(addedShape('new-a'), addedShape('new-b'), addedShape('new-c'));

    const ids = idsIn(serializeSlide(slide).xml);
    expect(new Set(ids).size).toBe(ids.length);
    expect(ids).not.toContain('0');
  });

  it('leaves untouched shapes on their original ids', () => {
    const slide = slideWith(existingShape('2', 'A') + existingShape('7', 'B'));
    const { xml } = serializeSlide(slide);
    expect(idsIn(xml)).toEqual(['1', '2', '7']);
  });

  it('never reuses an id already present in the part', () => {
    const slide = slideWith(existingShape('2', 'A') + existingShape('3', 'B'));
    slide.nodes.push(addedShape('new-a'));
    const ids = idsIn(serializeSlide(slide).xml);
    expect(ids.filter((id) => id === '2').length).toBe(1);
    expect(ids.filter((id) => id === '3').length).toBe(1);
  });

  it('writes only positive integers, as the schema requires', () => {
    const slide = slideWith(existingShape('2', 'A'));
    slide.nodes.push(addedShape('new-a'), addedShape('0'), addedShape('-5'));
    for (const id of idsIn(serializeSlide(slide).xml)) {
      expect(Number.isInteger(Number(id))).toBe(true);
      expect(Number(id)).toBeGreaterThanOrEqual(1);
    }
  });

  it('updates the model so the host and the file agree on the id', () => {
    const slide = slideWith(existingShape('2', 'A'));
    const added = addedShape('new-a');
    slide.nodes.push(added);

    const { xml } = serializeSlide(slide);
    // The host keys its node map and selection by node.id; a file id it does not know
    // about would leave the shape unreachable exactly as before.
    expect(idsIn(xml)).toContain(added.id);
  });

  it('carries the node name into the markup', () => {
    const slide = slideWith(existingShape('2', 'A'));
    const added = addedShape('new-a');
    added.name = 'My Shape';
    slide.nodes.push(added);
    expect(serializeSlide(slide).xml).toContain('name="My Shape"');
  });

  it('keeps added shapes in the presentationml namespace', () => {
    const slide = slideWith(existingShape('2', 'A'));
    const added = addedShape('new-a');
    slide.nodes.push(added);

    const { xml } = serializeSlide(slide);
    // An unnamespaced <sp> still renders here — the parser matches on localName — but it
    // is invalid OOXML, and PowerPoint drops the shape rather than complaining.
    expect(added.source.element?.namespaceURI).toBe(P);
    expect(xml).toContain('<p:sp>');
    expect(xml).not.toMatch(/<sp[ >]/);
    expect(xml).not.toMatch(/<cNvPr[ >]/);
  });
});
