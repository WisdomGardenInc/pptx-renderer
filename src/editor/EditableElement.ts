/** Framework-free editor values. Geometry and margins use 96-DPI pixels;
 * font sizes and point spacing use points, ratios use fractions. */
export interface EditableBullet {
  kind: 'char' | 'autonum';
  char?: string;
  font?: string;
  sizePct?: number;
  sizePt?: number;
  color?: string;
  numType?: string;
  startAt?: number;
}

export interface EditableTextRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontName?: string | null;
  fontEa?: string | null;
  fontCs?: string | null;
  fontSize?: number | null;
  color?: string | null;
  colorAlpha?: number | null;
  spc?: number | null;
  baseline?: number | null;
}

export interface EditableParagraph {
  runs: EditableTextRun[];
  align?: 'none' | 'left' | 'center' | 'right' | 'justify';
  lineSpacing?: number | null;
  lineSpacingIsPt?: boolean;
  spaceBefore?: number | null;
  spaceAfter?: number | null;
  indentLevel?: number;
  marL?: number | null;
  indent?: number | null;
  bullet?: EditableBullet | null;
}

export interface EditableElement {
  /** Opaque identity supplied by the embedding editor. */
  id: string;
  type: 'text' | 'shape' | 'image' | 'line' | 'table' | 'chart' | 'group' | 'smartart';
  nodeId?: string;
  sourceXml?: string;
  sourcePart?: string;
  relationships?: Record<string, { type: string; target: string; targetMode?: string }>;
  mediaPath?: string;
  /** Opaque image identity; resolving and uploading it belongs to the host. */
  imageKey?: string;
  x: number;
  y: number;
  width: number;
  height: number;
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  locked?: boolean;
  paragraphs?: EditableParagraph[];
  shapeType?: string | null;
  fillColor?: string | null;
  fillAlpha?: number | null;
  strokeColor?: string | null;
  strokeAlpha?: number | null;
  strokeWidth?: number | null;
  strokeDash?: 'solid' | 'dashed' | 'dotted' | 'dashdot';
  wrap?: string | null;
  vert?: string | null;
  vAlign?: 'top' | 'middle' | 'bottom';
  bodyPadding?: { l: number; r: number; t: number; b: number } | null;
  fontScale?: number | null;
  opacity?: number;
  picCrop?: { l?: number; r?: number; t?: number; b?: number } | null;
  zOrder?: number;
  isPlaceholder?: boolean;
  placeholderType?: string;
}
