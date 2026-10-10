import {
  parseXml,
  parseShapeNode,
  type ShapeNodeData,
  type SlideData,
} from '@wisdomgarden/pptx-renderer';

const DML = 'http://schemas.openxmlformats.org/drawingml/2006/main';
// Seed markup must declare presentationml and prefix its own elements with it. The
// renderer matches on localName and so draws an unnamespaced <sp> quite happily, but
// the file is then invalid OOXML and PowerPoint drops the shape — a divergence that
// only shows up once the deck leaves the editor.
const PML = 'http://schemas.openxmlformats.org/presentationml/2006/main';

/**
 * An id for a new node: the lowest positive integer the slide does not already use.
 *
 * It has to be a number. `cNvPr/@id` is `xsd:unsignedInt` and unique within the slide,
 * so the editor cannot invent an id shaped however it likes and expect it to survive a
 * save — an earlier `new-…` scheme wrote every added shape out under the seed markup's
 * own id, and on reload they collided and all but the last became unselectable.
 *
 * The whole part is scanned rather than just `slide.nodes`: group children and anything
 * the parser skipped still occupy ids. 1 is left alone, being conventionally the shape
 * tree's own `cNvPr`.
 */
export function freshId(slide: SlideData): string {
  const used = new Set<number>();

  const root = slide.root?.element;
  if (root) {
    const all = root.getElementsByTagName('*');
    for (let i = 0; i < all.length; i++) {
      if (all[i].localName !== 'cNvPr') continue;
      const value = Number(all[i].getAttribute('id'));
      if (Number.isInteger(value) && value >= 1) used.add(value);
    }
  }
  // Nodes added earlier in this session are not in the markup yet.
  for (const node of slide.nodes) {
    const value = Number(node.id);
    if (Number.isInteger(value) && value >= 1) used.add(value);
  }

  let id = 2;
  while (used.has(id)) id += 1;
  return String(id);
}

/**
 * Parse a seed `<sp>` OOXML string through the real shape parser, then stamp identity and
 * geometry. Using the library parser (not a hand-built object) means the new element renders
 * and restyles through exactly the same path as parsed elements.
 */
function buildShape(
  xml: string,
  opts: { id: string; name: string; x: number; y: number; w: number; h: number },
): ShapeNodeData {
  const node = parseShapeNode(parseXml(xml));
  node.id = opts.id;
  node.name = opts.name;
  node.position = { x: opts.x, y: opts.y };
  node.size = { w: opts.w, h: opts.h };
  node.rotation = 0;
  node.flipH = false;
  node.flipV = false;
  return node;
}

export function createTextBox(
  slide: SlideData,
  box: { x: number; y: number; w: number; h: number },
  text = 'Text',
): ShapeNodeData {
  const safe = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const xml = `
    <p:sp xmlns:a="${DML}" xmlns:p="${PML}">
      <p:nvSpPr><p:cNvPr id="0" name="Text Box"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="457200"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
      </p:spPr>
      <p:txBody>
        <a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr>
        <a:lstStyle/>
        <a:p>
          <a:r>
            <a:rPr lang="en-US" sz="1800" dirty="0">
              <a:solidFill><a:srgbClr val="000000"/></a:solidFill>
            </a:rPr>
            <a:t>${safe}</a:t>
          </a:r>
        </a:p>
      </p:txBody>
    </p:sp>`;
  return buildShape(xml, { id: freshId(slide), name: 'Text Box', ...box });
}

export function createRectangle(
  slide: SlideData,
  box: { x: number; y: number; w: number; h: number },
  fillHex = '4472C4',
): ShapeNodeData {
  const xml = `
    <p:sp xmlns:a="${DML}" xmlns:p="${PML}">
      <p:nvSpPr><p:cNvPr id="0" name="Rectangle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="${fillHex}"/></a:solidFill>
        <a:ln w="12700"><a:solidFill><a:srgbClr val="2E4E8F"/></a:solidFill></a:ln>
      </p:spPr>
      <p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p/></p:txBody>
    </p:sp>`;
  return buildShape(xml, { id: freshId(slide), name: 'Rectangle', ...box });
}

// ---- z-order over the typed slide.nodes array (later index = front) ---------

export type ReorderDir = 'front' | 'back' | 'forward' | 'backward';

export function reorderNode(slide: SlideData, id: string, dir: ReorderDir): void {
  const i = slide.nodes.findIndex((n) => n.id === id);
  if (i < 0) return;
  const [node] = slide.nodes.splice(i, 1);
  switch (dir) {
    case 'front':
      slide.nodes.push(node);
      break;
    case 'back':
      slide.nodes.unshift(node);
      break;
    case 'forward':
      slide.nodes.splice(Math.min(i + 1, slide.nodes.length), 0, node);
      break;
    case 'backward':
      slide.nodes.splice(Math.max(i - 1, 0), 0, node);
      break;
  }
}

export function deleteNode(slide: SlideData, id: string): boolean {
  const i = slide.nodes.findIndex((n) => n.id === id);
  if (i < 0) return false;
  slide.nodes.splice(i, 1);
  return true;
}
