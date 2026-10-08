/** Incrementally edit projected elements while retaining unmodified OOXML. */
import { parseXml } from '../parser/XmlParser';
import type { SlideData } from '../model/Slide';
import type {
  EditableElement,
  EditableParagraph,
  EditableTextRun,
} from '../editor/EditableElement';
import { resolvePartPath, sourceElement } from '../editor/parts';
import { A, P, R, SP_PR_ORDER, RPR_CHILD_ORDER, ensureChild } from './xmlEdit';

const PACKAGE_R = 'http://schemas.openxmlformats.org/package/2006/relationships';
const ORDER: Record<string, readonly string[]> = {
  spPr: SP_PR_ORDER,
  xfrm: ['off', 'ext', 'chOff', 'chExt'],
  pPr: [
    'lnSpc',
    'spcBef',
    'spcAft',
    'buClrTx',
    'buClr',
    'buSzTx',
    'buSzPct',
    'buSzPts',
    'buFontTx',
    'buFont',
    'buNone',
    'buAutoNum',
    'buChar',
    'buBlip',
    'tabLst',
    'defRPr',
    'extLst',
  ],
  rPr: RPR_CHILD_ORDER,
  txBody: ['bodyPr', 'lstStyle', 'p'],
  p: ['pPr', 'r', 'br', 'fld', 'tab', 'endParaRPr'],
};

function parseElement(xml: string): Element {
  const node = parseXml(xml);
  if (!node.exists()) throw new Error('Invalid editor XML.');
  return node.element!;
}

function child(parent: Element, name: string): Element | undefined {
  return Array.from(parent.children).find((element) => element.localName === name);
}

function ensure(parent: Element, name: string, ns = A): Element {
  return ensureChild(parent, ns, `${ns === P ? 'p' : 'a'}:${name}`, ORDER[parent.localName] ?? []);
}

function setAttribute(
  element: Element,
  name: string,
  value: string | number | null | undefined,
): void {
  if (value == null) element.removeAttribute(name);
  else element.setAttribute(name, String(value));
}

function nodeId(element: Element): string | null {
  return element.getElementsByTagNameNS(P, 'cNvPr')[0]?.getAttribute('id') ?? null;
}

function setColor(parent: Element, color: string | null | undefined, alpha?: number | null): void {
  for (const element of Array.from(parent.children)) {
    if (
      ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'].includes(
        element.localName,
      )
    )
      element.remove();
  }
  const fill = ensure(parent, color ? 'solidFill' : 'noFill');
  if (color) {
    const rgb = ensure(fill, 'srgbClr');
    rgb.setAttribute('val', color.replace(/^#/, ''));
    if (alpha != null && alpha < 1)
      ensure(rgb, 'alpha').setAttribute('val', String(Math.round(alpha * 100000)));
  }
}

function patchGeometry(source: Element, element: EditableElement): void {
  const container = child(source, 'spPr') ?? child(source, 'grpSpPr') ?? source;
  const transform = ensure(container, 'xfrm', container === source ? P : A);
  const off = ensure(transform, 'off');
  const ext = ensure(transform, 'ext');
  off.setAttribute('x', String(Math.round(element.x * 9525)));
  off.setAttribute('y', String(Math.round(element.y * 9525)));
  {
    ext.setAttribute('cx', String(Math.round(element.width * 9525)));
    ext.setAttribute('cy', String(Math.round(element.height * 9525)));
  }
  const rotation = 'rotation' in element ? element.rotation : undefined;
  setAttribute(transform, 'rot', rotation ? Math.round(rotation * 60000) : undefined);
  if ('flipH' in element) setAttribute(transform, 'flipH', element.flipH ? '1' : undefined);
  if ('flipV' in element) setAttribute(transform, 'flipV', element.flipV ? '1' : undefined);
}

function applyRunProperties(
  properties: Element,
  run: EditableTextRun,
  previous?: EditableTextRun,
): void {
  const attrs = {
    bold: 'b',
    italic: 'i',
    underline: 'u',
    fontSize: 'sz',
    spc: 'spc',
    baseline: 'baseline',
  } as const;
  for (const [key, attr] of Object.entries(attrs)) {
    const value = run[key as keyof typeof attrs];
    if (previous && value === previous[key as keyof typeof attrs]) continue;
    const xmlValue =
      key === 'fontSize'
        ? value == null
          ? undefined
          : Math.round(Number(value) * 100)
        : key === 'underline'
          ? value
            ? 'sng'
            : 'none'
          : key === 'bold' || key === 'italic'
            ? value
              ? '1'
              : '0'
            : (value as number | null | undefined);
    setAttribute(properties, attr, xmlValue);
  }
  for (const [key, script] of [
    ['fontName', 'latin'],
    ['fontEa', 'ea'],
    ['fontCs', 'cs'],
  ] as const) {
    if (previous && run[key] === previous[key]) continue;
    if (run[key]) ensure(properties, script).setAttribute('typeface', run[key]!);
  }
  if (!previous || run.color !== previous.color || run.colorAlpha !== previous.colorAlpha) {
    if (run.color) setColor(properties, run.color, run.colorAlpha);
  }
}

function applyParagraphProperties(
  properties: Element,
  paragraph: EditableParagraph,
  previous: EditableParagraph | undefined,
): void {
  if (!previous || JSON.stringify(paragraph.bullet) !== JSON.stringify(previous.bullet)) {
    for (const element of Array.from(properties.children)) {
      if (
        /^bu(?:ClrTx|Clr|SzTx|SzPct|SzPts|FontTx|Font|None|AutoNum|Char|Blip)$/.test(
          element.localName,
        )
      )
        element.remove();
    }
    const bullet = paragraph.bullet;
    if (!bullet) ensure(properties, 'buNone');
    else {
      if (bullet.color)
        ensure(ensure(properties, 'buClr'), 'srgbClr').setAttribute(
          'val',
          bullet.color.replace(/^#/, ''),
        );
      else ensure(properties, 'buClrTx');
      if (bullet.sizePt != null)
        ensure(properties, 'buSzPts').setAttribute('val', String(Math.round(bullet.sizePt * 100)));
      else if (bullet.sizePct != null)
        ensure(properties, 'buSzPct').setAttribute(
          'val',
          String(Math.round(bullet.sizePct * 100000)),
        );
      else ensure(properties, 'buSzTx');
      if (bullet.font) ensure(properties, 'buFont').setAttribute('typeface', bullet.font);
      else ensure(properties, 'buFontTx');
      if (bullet.kind === 'autonum') {
        const numbering = ensure(properties, 'buAutoNum');
        numbering.setAttribute('type', bullet.numType ?? 'arabicPeriod');
        setAttribute(numbering, 'startAt', bullet.startAt);
      } else ensure(properties, 'buChar').setAttribute('char', bullet.char ?? '•');
    }
  }
  if (!previous || paragraph.align !== previous.align) {
    setAttribute(
      properties,
      'algn',
      ({ left: 'l', center: 'ctr', right: 'r', justify: 'just', none: undefined } as const)[
        paragraph.align ?? 'none'
      ],
    );
  }
  for (const [key, attr] of [
    ['marL', 'marL'],
    ['indent', 'indent'],
    ['indentLevel', 'lvl'],
  ] as const) {
    if (previous && paragraph[key] === previous[key]) continue;
    const value = paragraph[key];
    setAttribute(
      properties,
      attr,
      value == null
        ? key === 'indentLevel'
          ? undefined
          : 0
        : key === 'indentLevel'
          ? value
          : Math.round(value * 9525),
    );
  }
  for (const [key, name] of [
    ['lineSpacing', 'lnSpc'],
    ['spaceBefore', 'spcBef'],
    ['spaceAfter', 'spcAft'],
  ] as const) {
    if (
      previous &&
      paragraph[key] === previous[key] &&
      (key !== 'lineSpacing' || paragraph.lineSpacingIsPt === previous.lineSpacingIsPt)
    )
      continue;
    child(properties, name)?.remove();
    const value = paragraph[key];
    if (value != null) {
      const spacing = ensure(properties, name);
      const points = key !== 'lineSpacing' || paragraph.lineSpacingIsPt;
      ensure(spacing, points ? 'spcPts' : 'spcPct').setAttribute(
        'val',
        String(Math.round(value * (points ? 100 : 100000))),
      );
    }
  }
}

function patchParagraphs(
  source: Element,
  paragraphs: EditableParagraph[],
  previous: EditableParagraph[] | undefined,
): void {
  const body = ensure(source, 'txBody', P);
  ensure(body, 'bodyPr');
  ensure(body, 'lstStyle');
  const oldParagraphs = Array.from(body.children).filter((element) => element.localName === 'p');
  if (
    previous &&
    oldParagraphs.length === paragraphs.length &&
    paragraphs.every(
      (paragraph, index) =>
        JSON.stringify(paragraph.runs) === JSON.stringify(previous[index]?.runs),
    )
  ) {
    // Paragraph formatting such as alignment does not require recreating text runs.
    for (const [index, paragraph] of paragraphs.entries()) {
      applyParagraphProperties(ensure(oldParagraphs[index], 'pPr'), paragraph, previous[index]);
    }
    return;
  }
  for (const old of oldParagraphs) old.remove();
  for (const [index, paragraph] of paragraphs.entries()) {
    const old = oldParagraphs[index];
    const paragraphElement = old
      ? (old.cloneNode(true) as Element)
      : source.ownerDocument.createElementNS(A, 'a:p');
    applyParagraphProperties(ensure(paragraphElement, 'pPr'), paragraph, previous?.[index]);
    const runs = Array.from(paragraphElement.children).filter((element) =>
      ['r', 'br', 'fld', 'tab'].includes(element.localName),
    );
    for (const run of runs) run.remove();
    for (const [runIndex, run] of paragraph.runs.entries()) {
      const oldRun = runs[runIndex];
      const parts = run.text.split(/(\n|\t)/).filter(Boolean);
      for (const text of parts) {
        const runElement =
          oldRun && text !== '\n' && text !== '\t' && ['r', 'fld'].includes(oldRun.localName)
            ? (oldRun.cloneNode(true) as Element)
            : source.ownerDocument.createElementNS(
                A,
                text === '\n' ? 'a:br' : text === '\t' ? 'a:tab' : 'a:r',
              );
        if (text !== '\t')
          applyRunProperties(ensure(runElement, 'rPr'), run, previous?.[index]?.runs[runIndex]);
        if (text !== '\n' && text !== '\t') ensure(runElement, 't').textContent = text;
        paragraphElement.insertBefore(runElement, child(paragraphElement, 'endParaRPr') ?? null);
      }
    }
    body.appendChild(paragraphElement);
  }
}

function patchStyle(
  source: Element,
  element: EditableElement,
  previous: EditableElement | undefined,
): void {
  const spPr = child(source, 'spPr');
  if (
    spPr &&
    'fillColor' in element &&
    (!previous ||
      !('fillColor' in previous) ||
      element.fillColor !== previous.fillColor ||
      element.fillAlpha !== previous.fillAlpha)
  ) {
    setColor(spPr, element.fillColor, element.fillAlpha);
  }
  if (spPr && 'strokeColor' in element) {
    const old = previous && 'strokeColor' in previous ? previous : undefined;
    if (!old || element.strokeColor !== old.strokeColor || element.strokeAlpha !== old.strokeAlpha)
      setColor(ensure(spPr, 'ln'), element.strokeColor, element.strokeAlpha);
    if (!old || element.strokeWidth !== old.strokeWidth)
      setAttribute(
        ensure(spPr, 'ln'),
        'w',
        element.strokeWidth == null ? undefined : Math.round(element.strokeWidth * 9525),
      );
    if (!old || element.strokeDash !== old.strokeDash) {
      if (element.strokeDash)
        ensure(ensure(spPr, 'ln'), 'prstDash').setAttribute(
          'val',
          ({ solid: 'solid', dashed: 'dash', dotted: 'dot', dashdot: 'dashDot' } as const)[
            element.strokeDash
          ],
        );
    }
  }
  if (element.type === 'text' || element.type === 'shape') {
    const old = previous?.type === 'text' || previous?.type === 'shape' ? previous : undefined;
    if (JSON.stringify(element.paragraphs) !== JSON.stringify(old?.paragraphs))
      patchParagraphs(source, element.paragraphs ?? [], old?.paragraphs);
    const body = child(source, 'txBody');
    if (body) {
      const properties = ensure(body, 'bodyPr');
      if (!old || element.wrap !== old.wrap) setAttribute(properties, 'wrap', element.wrap);
      if (!old || element.vAlign !== old.vAlign)
        setAttribute(
          properties,
          'anchor',
          ({ top: 't', middle: 'ctr', bottom: 'b' } as const)[element.vAlign ?? 'top'],
        );
      if (!old || JSON.stringify(element.bodyPadding) !== JSON.stringify(old.bodyPadding)) {
        for (const [key, attr] of [
          ['l', 'lIns'],
          ['r', 'rIns'],
          ['t', 'tIns'],
          ['b', 'bIns'],
        ] as const)
          setAttribute(
            properties,
            attr,
            element.bodyPadding ? Math.round(element.bodyPadding[key] * 9525) : undefined,
          );
      }
    }
  }
  if (
    element.type === 'image' &&
    element.opacity !== (previous?.type === 'image' ? previous.opacity : undefined)
  ) {
    const blip = child(child(source, 'blipFill') ?? source, 'blip');
    if (blip) {
      for (const modifier of Array.from(blip.children)) {
        if (['alphaModFix', 'alphaMod', 'alphaOff'].includes(modifier.localName)) modifier.remove();
      }
      const opacity = Math.max(0, Math.min(1, element.opacity ?? 1));
      if (opacity < 1) {
        const modifier = source.ownerDocument.createElementNS(A, 'a:alphaModFix');
        modifier.setAttribute('amt', String(Math.round(opacity * 100000)));
        blip.insertBefore(modifier, child(blip, 'extLst') ?? null);
      }
    }
  }
  if (
    element.type === 'image' &&
    JSON.stringify(element.picCrop) !==
      JSON.stringify(previous?.type === 'image' ? previous.picCrop : undefined)
  ) {
    const fill = child(source, 'blipFill');
    if (fill) {
      child(fill, 'srcRect')?.remove();
      if (element.picCrop) {
        const crop = source.ownerDocument.createElementNS(A, 'a:srcRect');
        for (const key of ['l', 'r', 't', 'b'] as const)
          setAttribute(crop, key, Math.round((element.picCrop[key] ?? 0) * 100000));
        fill.insertBefore(crop, child(fill, 'stretch') ?? child(fill, 'tile') ?? null);
      }
    }
  }
}

function newShape(document: Document, element: EditableElement, id: string): Element {
  const isImage = element.type === 'image';
  const source = document.createElementNS(P, isImage ? 'p:pic' : 'p:sp');
  const nv = ensure(source, isImage ? 'nvPicPr' : 'nvSpPr', P);
  const properties = ensure(nv, 'cNvPr', P);
  properties.setAttribute('id', id);
  properties.setAttribute('name', isImage ? 'Picture' : 'Text');
  const cNv = ensure(nv, isImage ? 'cNvPicPr' : 'cNvSpPr', P);
  if (element.type === 'text') cNv.setAttribute('txBox', '1');
  ensure(nv, 'nvPr', P);
  if (isImage) {
    const fill = ensure(source, 'blipFill', P);
    ensure(fill, 'blip');
    ensure(ensure(fill, 'stretch'), 'fillRect');
  }
  const propertiesElement = ensure(source, 'spPr', P);
  if (element.type === 'text') {
    setColor(propertiesElement, null);
    setColor(ensure(propertiesElement, 'ln'), null);
  }
  const geometry = ensure(propertiesElement, 'prstGeom');
  geometry.setAttribute('prst', 'shapeType' in element ? (element.shapeType ?? 'rect') : 'rect');
  ensure(geometry, 'avLst');
  return source;
}

function addRelationship(root: Element, target: string, type: string, targetMode?: string): string {
  const existing = Array.from(root.children).find(
    (rel) =>
      rel.getAttribute('Target') === target &&
      rel.getAttribute('Type') === type &&
      (rel.getAttribute('TargetMode') ?? undefined) === targetMode,
  );
  if (existing) return existing.getAttribute('Id')!;
  const used = new Set(Array.from(root.children).map((rel) => rel.getAttribute('Id')));
  let id = 'rIdEditor1';
  for (let counter = 2; used.has(id); counter++) id = `rIdEditor${counter}`;
  const rel = root.ownerDocument.createElementNS(PACKAGE_R, 'Relationship');
  rel.setAttribute('Id', id);
  rel.setAttribute('Type', type);
  rel.setAttribute('Target', target);
  if (targetMode) rel.setAttribute('TargetMode', targetMode);
  root.appendChild(rel);
  return id;
}

function relativeTarget(source: string, target: string): string {
  const base = source.split('/').slice(0, -1);
  const path = target.split('/');
  while (base.length && path.length && base[0] === path[0]) {
    base.shift();
    path.shift();
  }
  return [...base.map(() => '..'), ...path].join('/');
}

function remapRelationships(
  source: Element,
  element: EditableElement,
  slide: SlideData,
  rels: Element,
): void {
  if (!element.sourcePart || element.sourcePart === slide.slidePath) return;
  for (const node of [source, ...Array.from(source.getElementsByTagName('*'))]) {
    for (const attribute of Array.from(node.attributes)) {
      if (attribute.namespaceURI !== R) continue;
      const relationship = element.relationships?.[attribute.value];
      if (!relationship) throw new Error('The copied object has an unavailable PPTX relationship.');
      const target =
        relationship.targetMode === 'External'
          ? relationship.target
          : relativeTarget(
              slide.slidePath,
              resolvePartPath(element.sourcePart, relationship.target),
            );
      attribute.value = addRelationship(rels, target, relationship.type, relationship.targetMode);
    }
  }
}

export function editSlideElements(
  slide: SlideData,
  elements: readonly EditableElement[],
  previous: readonly EditableElement[],
  relationshipsXml: string | null,
  imageParts: Map<string, string>,
): { xml: string; relsXml: string | null; ids: Map<string, string> } {
  const root = slide.root?.exists() ? slide.root.element : undefined;
  const treeNode = slide.root?.child('cSld').child('spTree');
  const tree = treeNode?.exists() ? treeNode.element : undefined;
  if (!root || !tree) throw new Error('Missing PPTX slide tree.');
  const rels = parseElement(relationshipsXml ?? `<Relationships xmlns="${PACKAGE_R}"/>`);
  const previousById = new Map(previous.map((element) => [element.id, element]));
  const content = elements;
  const ids = new Map<string, string>();
  let nextId =
    Math.max(
      1,
      ...Array.from(root.getElementsByTagNameNS(P, 'cNvPr')).map(
        (node) => Number(node.getAttribute('id')) || 0,
      ),
    ) + 1;
  for (const old of previous) {
    if (old.nodeId && !content.some((element) => element.id === old.id))
      sourceElement(slide, old.nodeId)?.remove();
  }
  for (const element of content) {
    const old = previousById.get(element.id);
    let source = old?.nodeId ? sourceElement(slide, old.nodeId) : undefined;
    const created = !source;
    const cloned = created && !!element.sourceXml;
    const id = source ? nodeId(source)! : String(nextId++);
    if (!source) {
      source = element.sourceXml
        ? (root.ownerDocument.importNode(parseElement(element.sourceXml), true) as Element)
        : newShape(root.ownerDocument, element, id);
      const copiedIds = new Map<string, string>();
      Array.from(source.getElementsByTagNameNS(P, 'cNvPr')).forEach((properties, index) => {
        const replacement = index === 0 ? id : String(nextId++);
        copiedIds.set(properties.getAttribute('id') ?? '', replacement);
        properties.setAttribute('id', replacement);
      });
      for (const tag of ['stCxn', 'endCxn'])
        for (const connection of Array.from(source.getElementsByTagNameNS(A, tag))) {
          const replacement = copiedIds.get(connection.getAttribute('id') ?? '');
          if (replacement) connection.setAttribute('id', replacement);
        }
      remapRelationships(source, element, slide, rels);
    }
    ids.set(id, element.id);
    if (
      created ||
      JSON.stringify([
        element.x,
        element.y,
        'width' in element ? element.width : 0,
        'height' in element ? element.height : 0,
        'rotation' in element ? element.rotation : 0,
        'flipH' in element ? element.flipH : false,
        'flipV' in element ? element.flipV : false,
      ]) !==
        JSON.stringify([
          old?.x,
          old?.y,
          old && 'width' in old ? old.width : 0,
          old && 'height' in old ? old.height : 0,
          old && 'rotation' in old ? old.rotation : 0,
          old && 'flipH' in old ? old.flipH : false,
          old && 'flipV' in old ? old.flipV : false,
        ])
    )
      patchGeometry(source, element);
    patchStyle(source, element, cloned ? element : old);
    if (element.locked !== old?.locked) {
      const names = (
        {
          sp: ['nvSpPr', 'cNvSpPr', 'spLocks'],
          pic: ['nvPicPr', 'cNvPicPr', 'picLocks'],
          grpSp: ['nvGrpSpPr', 'cNvGrpSpPr', 'grpSpLocks'],
          graphicFrame: ['nvGraphicFramePr', 'cNvGraphicFramePr', 'graphicFrameLocks'],
        } as Record<string, string[]>
      )[source.localName];
      if (names) {
        const locks = ensure(ensure(ensure(source, names[0], P), names[1], P), names[2]);
        for (const name of ['noMove', 'noResize']) {
          if (element.locked) locks.setAttribute(name, '1');
          else locks.removeAttribute(name);
        }
      }
    }
    if (
      element.type === 'image' &&
      (!old || old.type !== 'image' || element.imageKey !== old.imageKey)
    ) {
      const path = imageParts.get(element.id) ?? element.mediaPath;
      if (!path) throw new Error('The image has not been embedded into the PPTX.');
      const id = addRelationship(rels, relativeTarget(slide.slidePath, path), `${R}/image`);
      ensure(ensure(source, 'blipFill', P), 'blip').setAttributeNS(R, 'r:embed', id);
    }
    tree.insertBefore(source, child(tree, 'extLst') ?? null);
  }
  return {
    xml: new XMLSerializer().serializeToString(root),
    relsXml: new XMLSerializer().serializeToString(rels),
    ids,
  };
}
