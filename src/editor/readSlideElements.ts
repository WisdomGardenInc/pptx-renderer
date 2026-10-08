/** Read framework-free editor values from a materialized slide. */
import type { PresentationData } from '../model/Presentation';
import type { SafeXmlNode } from '../parser/XmlParser';
import type { ShapeNodeData } from '../model/nodes/ShapeNode';
import type { SlideData, SlideNode } from '../model/Slide';
import type { ThemeData } from '../model/Theme';
import type {
  EditableBullet,
  EditableParagraph,
  EditableTextRun,
  EditableElement,
} from './EditableElement';
import { sourceElement, resolvePartPath } from './parts';

function slideTheme(presentation: PresentationData, slide: SlideData): ThemeData | undefined {
  const layout = presentation.slideToLayout.get(slide.index);
  const master = layout && presentation.layoutToMaster.get(layout);
  const theme = master && presentation.masterToTheme.get(master);
  return theme ? presentation.themes.get(theme) : undefined;
}

function readColor(fill: SafeXmlNode | undefined, theme?: ThemeData): string | null {
  if (!fill?.exists() || fill.localName === 'noFill') return null;
  const color = fill
    .allChildren()
    .find((child) => ['srgbClr', 'schemeClr', 'sysClr'].includes(child.localName));
  if (!color) return null;
  const slot = color.attr('val') ?? '';
  const value =
    color.localName === 'schemeClr'
      ? theme?.colorScheme.get(
          ({ tx1: 'dk1', tx2: 'dk2', bg1: 'lt1', bg2: 'lt2' } as Record<string, string>)[slot] ??
            slot,
        )
      : (color.attr('lastClr') ?? slot);
  return value && /^[a-f\d]{6}$/i.test(value) ? `#${value}` : null;
}

function property(sources: (SafeXmlNode | undefined)[], name: string): string | undefined {
  return sources.map((source) => source?.attr(name)).find((value) => value !== undefined);
}

function fontName(
  sources: (SafeXmlNode | undefined)[],
  script: string,
  theme?: ThemeData,
): string | undefined {
  const font = sources.map((source) => source?.child(script).attr('typeface')).find(Boolean);
  if (font?.startsWith('+mj-'))
    return theme?.majorFont[script as 'latin' | 'ea' | 'cs'] || undefined;
  if (font?.startsWith('+mn-'))
    return theme?.minorFont[script as 'latin' | 'ea' | 'cs'] || undefined;
  return font || theme?.minorFont[script as 'latin' | 'ea' | 'cs'] || undefined;
}

function textRun(
  text: string,
  sources: (SafeXmlNode | undefined)[],
  theme?: ThemeData,
): EditableTextRun {
  const size = property(sources, 'sz');
  const color = sources.map((source) => readColor(source?.child('solidFill'), theme)).find(Boolean);
  return {
    text,
    bold: property(sources, 'b') === '1',
    italic: property(sources, 'i') === '1',
    underline: !['none', undefined].includes(property(sources, 'u')),
    fontSize: size ? Number(size) / 100 : 18,
    fontName: fontName(sources, 'latin', theme),
    fontEa: fontName(sources, 'ea', theme),
    fontCs: fontName(sources, 'cs', theme),
    color: color ?? '#000000',
  };
}

function levelStyle(style: SafeXmlNode | undefined, level: string): SafeXmlNode | undefined {
  return style?.child(level).exists() ? style.child(level) : style?.child('defPPr');
}

function placeholderProperties(node: SafeXmlNode): SafeXmlNode {
  const shape = node.child('nvSpPr').child('nvPr').child('ph');
  return shape.exists() ? shape : node.child('nvPicPr').child('nvPr').child('ph');
}

function readBullet(
  sources: (SafeXmlNode | undefined)[],
  theme?: ThemeData,
): EditableBullet | null {
  const props = sources.filter((source): source is SafeXmlNode => !!source?.exists());
  const kind = props.find((source) =>
    ['buNone', 'buChar', 'buAutoNum', 'buBlip'].some((name) => source.child(name).exists()),
  );
  if (!kind || kind.child('buNone').exists() || kind.child('buBlip').exists()) return null;
  const numbering = kind.child('buAutoNum');
  const bullet: EditableBullet = numbering.exists()
    ? {
        kind: 'autonum',
        numType: numbering.attr('type') ?? 'arabicPeriod',
        startAt: numbering.numAttr('startAt'),
      }
    : { kind: 'char', char: kind.child('buChar').attr('char') ?? '•' };
  const font = props.find(
    (source) => source.child('buFont').exists() || source.child('buFontTx').exists(),
  );
  if (font?.child('buFont').exists()) bullet.font = font.child('buFont').attr('typeface');
  const size = props.find((source) =>
    ['buSzTx', 'buSzPct', 'buSzPts'].some((name) => source.child(name).exists()),
  );
  const percent = size?.child('buSzPct').numAttr('val');
  const points = size?.child('buSzPts').numAttr('val');
  if (percent != null) bullet.sizePct = percent / 100000;
  if (points != null) bullet.sizePt = points / 100;
  const color = props.find(
    (source) => source.child('buClr').exists() || source.child('buClrTx').exists(),
  );
  if (color?.child('buClr').exists())
    bullet.color = readColor(color.child('buClr'), theme) ?? undefined;
  return bullet;
}

function pictureOpacity(blip: SafeXmlNode): number {
  const fixed = (blip.child('alphaModFix').numAttr('amt') ?? 100000) / 100000;
  const mod = (blip.child('alphaMod').numAttr('val') ?? 100000) / 100000;
  const offset = (blip.child('alphaOff').numAttr('val') ?? 0) / 100000;
  return Math.max(0, Math.min(1, fixed * mod + offset));
}

function paragraphs(
  node: ShapeNodeData,
  presentation: PresentationData,
  slide: SlideData,
): EditableParagraph[] {
  const theme = slideTheme(presentation, slide);
  const layout = presentation.slideToLayout.get(slide.index);
  const masterPath = layout && presentation.layoutToMaster.get(layout);
  const master = masterPath ? presentation.masters.get(masterPath) : undefined;
  const layoutData = layout ? presentation.layouts.get(layout) : undefined;
  const layoutPlaceholder = node.placeholder
    ? layoutData?.placeholders.find(
        (entry) =>
          (placeholderProperties(entry.node).numAttr('idx') ?? 0) === (node.placeholder?.idx ?? 0),
      )?.node
    : undefined;
  const placeholderType = layoutPlaceholder
    ? placeholderProperties(layoutPlaceholder).attr('type')
    : node.placeholder?.type;
  const masterType =
    placeholderType === 'ctrTitle'
      ? 'title'
      : ['obj', 'subTitle', 'pic', 'chart', 'clipArt', 'dgm', 'media', 'tbl'].includes(
            placeholderType ?? 'obj',
          )
        ? 'body'
        : placeholderType;
  const masterPlaceholder = node.placeholder
    ? master?.placeholders.find(
        (source) => (placeholderProperties(source).attr('type') ?? 'obj') === masterType,
      )
    : undefined;
  const style =
    master?.textStyles[
      masterType === 'title'
        ? 'titleStyle'
        : node.placeholder && ['body', 'dt', 'ftr', 'sldNum'].includes(masterType ?? '')
          ? 'bodyStyle'
          : 'otherStyle'
    ];
  return (node.textBody?.paragraphs ?? []).map((paragraph) => {
    const props = paragraph.properties;
    const level = `lvl${paragraph.level + 1}pPr`;
    const sources = [
      props,
      levelStyle(node.textBody?.listStyle, level),
      levelStyle(layoutPlaceholder?.child('txBody').child('lstStyle'), level),
      levelStyle(masterPlaceholder?.child('txBody').child('lstStyle'), level),
      levelStyle(style, level),
      levelStyle(master?.defaultTextStyle, level),
      levelStyle(presentation.defaultTextStyle, level),
    ];
    const defaults = sources.map((source) => source?.child('defRPr'));
    const align = property(sources, 'algn');
    const margin = property(sources, 'marL');
    const indent = property(sources, 'indent');
    const paragraphChild = (name: string) =>
      sources.map((source) => source?.child(name)).find((child) => child?.exists());
    const spacing = paragraphChild('lnSpc');
    const before = paragraphChild('spcBef')?.child('spcPts').numAttr('val');
    const after = paragraphChild('spcAft')?.child('spcPts').numAttr('val');
    return {
      bullet: readBullet(sources, theme),
      runs: paragraph.runs.map((run) =>
        textRun(run.text, [run.properties, ...defaults, paragraph.endParaRPr], theme),
      ),
      align:
        ({ l: 'left', ctr: 'center', r: 'right', just: 'justify' } as const)[align as 'l'] ??
        'left',
      indentLevel: paragraph.level,
      marL: margin != null ? Number(margin) / 9525 : null,
      indent: indent != null ? Number(indent) / 9525 : null,
      lineSpacing:
        spacing?.child('spcPts').numAttr('val') != null
          ? spacing.child('spcPts').numAttr('val')! / 100
          : spacing?.child('spcPct').numAttr('val') != null
            ? spacing.child('spcPct').numAttr('val')! / 100000
            : null,
      lineSpacingIsPt: spacing?.child('spcPts').exists() ?? false,
      spaceBefore: before != null ? before / 100 : null,
      spaceAfter: after != null ? after / 100 : null,
    };
  });
}

export function readSlideElements(
  presentation: PresentationData,
  slide: SlideData,
): EditableElement[] {
  const theme = slideTheme(presentation, slide);
  return slide.nodes.map((node, index): EditableElement => {
    const id = node.id;
    const lockNames = (
      {
        sp: ['nvSpPr', 'cNvSpPr', 'spLocks'],
        pic: ['nvPicPr', 'cNvPicPr', 'picLocks'],
        grpSp: ['nvGrpSpPr', 'cNvGrpSpPr', 'grpSpLocks'],
        graphicFrame: ['nvGraphicFramePr', 'cNvGraphicFramePr', 'graphicFrameLocks'],
      } as Record<string, string[]>
    )[(node.source.exists() ? node.source.element?.localName : undefined) ?? ''];
    const locks = lockNames
      ? node.source.child(lockNames[0]).child(lockNames[1]).child(lockNames[2])
      : null;
    const base = {
      id,
      nodeId: node.id,
      locked: locks?.attr('noMove') === '1' && locks?.attr('noResize') === '1',
      sourceXml: sourceElement(slide, node.id)
        ? new XMLSerializer().serializeToString(sourceElement(slide, node.id)!)
        : undefined,
      sourcePart: slide.slidePath,
      relationships: Object.fromEntries(slide.rels),
      x: node.position.x,
      y: node.position.y,
      width: node.size.w,
      height: node.size.h,
      rotation: node.rotation,
      flipH: node.flipH,
      flipV: node.flipV,
      zOrder: index,
      isPlaceholder: !!node.placeholder,
      placeholderType: node.placeholder?.type,
    };
    if (node.nodeType === 'shape') {
      const text = paragraphs(node, presentation, slide);
      const body = node.textBody?.bodyProperties;
      const fontScale = body?.child('normAutofit').numAttr('fontScale');
      return {
        ...base,
        type: node.source.child('nvSpPr').child('cNvSpPr').attr('txBox') === '1' ? 'text' : 'shape',
        paragraphs: text,
        shapeType: node.presetGeometry ?? 'rect',
        fillColor: readColor(node.fill, theme),
        strokeColor: readColor(node.line?.child('solidFill'), theme),
        strokeWidth: (node.line?.numAttr('w') ?? 0) / 9525,
        vAlign:
          ({ t: 'top', ctr: 'middle', b: 'bottom' } as const)[body?.attr('anchor') as 't'] ?? 'top',
        wrap: body?.attr('wrap') ?? node.textBody?.layoutBodyProperties?.attr('wrap'),
        vert: body?.attr('vert'),
        fontScale: fontScale != null ? fontScale / 100000 : null,
        bodyPadding: {
          l: (body?.numAttr('lIns') ?? 91440) / 9525,
          r: (body?.numAttr('rIns') ?? 91440) / 9525,
          t: (body?.numAttr('tIns') ?? 45720) / 9525,
          b: (body?.numAttr('bIns') ?? 45720) / 9525,
        },
      };
    }
    if (node.nodeType === 'picture') {
      const rel = slide.rels.get(node.blipEmbed ?? '');
      const path = rel ? resolvePartPath(slide.slidePath, rel.target) : '';
      return {
        ...base,
        type: 'image',
        opacity: pictureOpacity(node.source.child('blipFill').child('blip')),
        mediaPath: path,
        picCrop: node.crop
          ? { l: node.crop.left, r: node.crop.right, t: node.crop.top, b: node.crop.bottom }
          : null,
      };
    }
    return { ...base, type: node.nodeType };
  });
}

export function imagePartPaths(slide: SlideData): string[] {
  return slide.nodes.flatMap((node: SlideNode) => {
    if (node.nodeType !== 'picture') return [];
    const rel = slide.rels.get(node.blipEmbed ?? '');
    return rel ? [resolvePartPath(slide.slidePath, rel.target)] : [];
  });
}
