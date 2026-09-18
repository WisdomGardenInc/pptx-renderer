import {
  parseXml,
  parseShapeNode,
  type ShapeNodeData,
  type SlideData,
} from '@wisdomgarden/pptx-renderer';

const DML = 'http://schemas.openxmlformats.org/drawingml/2006/main';

let counter = 0;

/** A unique node id that cannot collide with PPTX numeric ids (avoids hit-test confusion). */
export function freshId(slide: SlideData): string {
  const existing = new Set(slide.nodes.map((n) => n.id));
  let id: string;
  do {
    id = `new-${Date.now().toString(36)}-${counter++}`;
  } while (existing.has(id));
  return id;
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
    <sp xmlns:a="${DML}">
      <nvSpPr><cNvPr id="0" name="Text Box"/><cNvSpPr txBox="1"/><nvPr/></nvSpPr>
      <spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="457200"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
      </spPr>
      <txBody>
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
      </txBody>
    </sp>`;
  return buildShape(xml, { id: freshId(slide), name: 'Text Box', ...box });
}

export function createRectangle(
  slide: SlideData,
  box: { x: number; y: number; w: number; h: number },
  fillHex = '4472C4',
): ShapeNodeData {
  const xml = `
    <sp xmlns:a="${DML}">
      <nvSpPr><cNvPr id="0" name="Rectangle"/><cNvSpPr/><nvPr/></nvSpPr>
      <spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="${fillHex}"/></a:solidFill>
        <a:ln w="12700"><a:solidFill><a:srgbClr val="2E4E8F"/></a:solidFill></a:ln>
      </spPr>
      <txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p/></txBody>
    </sp>`;
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
