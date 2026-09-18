import {
  parseZip,
  buildPresentation,
  renderSlide,
  type PresentationData,
  type SlideData,
  type SlideHandle,
  type BaseNodeData,
  type ShapeNodeData,
} from '@wisdomgarden/pptx-renderer';
import { TransformController, type Geometry } from './TransformController';
import {
  applyTextStyle as facadeApplyTextStyle,
  applyFillColor as facadeApplyFillColor,
  setPlainText,
  readPlainText,
  readTextStyle,
  type TextStylePatch,
} from './StyleFacade';
import {
  createTextBox,
  createRectangle,
  deleteNode,
  reorderNode,
  type ReorderDir,
} from './nodeOps';

/** A node plus the DOM element the renderer produced for it, for the current slide render. */
export interface RenderedNode {
  node: BaseNodeData;
  element: HTMLElement;
}

export interface RenderHostState {
  loaded: boolean;
  slideCount: number;
  slideIndex: number;
  /** Intrinsic slide size in px (the model coordinate space). */
  width: number;
  height: number;
  /** Current fit scale applied to the rendered slide. */
  scale: number;
  /** id of the selected top-level node, or null. */
  selectedId: string | null;
}

type Listener = (state: RenderHostState) => void;

/**
 * Framework-agnostic controller that owns a loaded presentation and renders one slide at a
 * time into a mount element. It builds a DOM↔model map via the library's `onNodeRendered`
 * hook, so a click on a rendered element resolves back to its model node. Selection is
 * kept at the top-level node granularity (groups are treated atomically for the POC).
 */
export class RenderHost {
  private presentation: PresentationData | null = null;
  private slideIndex = 0;
  private handle: SlideHandle | null = null;
  private scale = 1;
  private selectedId: string | null = null;

  /** All rendered nodes on the current slide, incl. group descendants, keyed by node id. */
  private nodeMap = new Map<string, RenderedNode>();
  /** ids of the current slide's top-level nodes (slide.nodes). */
  private topLevelIds = new Set<string>();
  /** Persistent media blob-URL cache so re-renders reuse (not revoke+refetch) image blobs. */
  private mediaCache = new Map<string, string>();

  private listeners = new Set<Listener>();

  // DOM structure: mount > stageWrap > [slideEl (scaled), overlayLayer]
  private stageWrap: HTMLElement;
  private overlayLayer: HTMLElement;
  private transform: TransformController;
  /** Element transform captured at the start of a live move, so previews compose cleanly. */
  private dragBase: string | null = null;
  /** Active inline text editor (double-click to edit), or null. */
  private textEditor: HTMLTextAreaElement | null = null;
  /** The shape currently being text-edited. Commit uses this, not the live selection,
   *  so clicking inside the editor (which can change selection) never loses the edit. */
  private editingShape: ShapeNodeData | null = null;

  constructor(private mount: HTMLElement) {
    this.mount.style.position = 'relative';
    this.mount.style.overflow = 'auto';

    this.stageWrap = document.createElement('div');
    this.stageWrap.style.position = 'relative';
    this.stageWrap.style.margin = '0 auto';

    this.overlayLayer = document.createElement('div');
    this.overlayLayer.style.position = 'absolute';
    this.overlayLayer.style.left = '0';
    this.overlayLayer.style.top = '0';
    this.overlayLayer.style.pointerEvents = 'none';
    this.overlayLayer.style.zIndex = '10';

    this.stageWrap.appendChild(this.overlayLayer);
    this.mount.appendChild(this.stageWrap);

    this.transform = new TransformController(this.overlayLayer, {
      getScale: () => this.scale,
      onPreview: (geom, kind) => this.previewTransform(geom, kind),
      onCommit: (geom) => this.commitTransform(geom),
      onBoxPointerDown: () => this.registerClick(this.selectedId),
    });

    this.stageWrap.addEventListener('pointerdown', this.handlePointerDown);
    document.addEventListener('keydown', this.handleKeyDown);
  }

  // Manual double-click detection. Native dblclick is unreliable here: the first click
  // selects a shape and overlays the transform box, so the two clicks have different
  // targets and the browser never fires dblclick. We correlate clicks by node id + time
  // across both the stage (unselected shape) and the transform box (selected shape).
  private lastClickId: string | null = null;
  private lastClickT = 0;

  private registerClick(id: string | null): boolean {
    const now = Date.now();
    if (id && id === this.lastClickId && now - this.lastClickT < 400) {
      this.lastClickId = null;
      this.lastClickT = 0;
      const rn = this.nodeMap.get(id);
      if (rn && rn.node.nodeType === 'shape') {
        this.startTextEdit(rn.node as ShapeNodeData);
        return true;
      }
      return false;
    }
    this.lastClickId = id;
    this.lastClickT = now;
    return false;
  }

  private handleKeyDown = (e: KeyboardEvent): void => {
    // Ignore when typing in a form control / editable region.
    const t = e.target as HTMLElement | null;
    const tag = t?.tagName;
    if (tag === 'INPUT' || tag === 'TEXTAREA' || t?.isContentEditable) return;
    if ((e.key === 'Delete' || e.key === 'Backspace') && this.selectedId) {
      e.preventDefault();
      this.deleteSelected();
    } else if (e.key === 'Escape') {
      this.select(null);
    }
  };

  private nodeGeometry(node: BaseNodeData): Geometry {
    return {
      x: node.position.x,
      y: node.position.y,
      w: node.size.w,
      h: node.size.h,
      rotation: node.rotation,
    };
  }

  /** Live preview during a drag: move translates the actual element; resize/rotate show on the box. */
  private previewTransform(geom: Geometry, kind: string): void {
    const rn = this.selectedId ? this.nodeMap.get(this.selectedId) : null;
    if (!rn) return;
    if (kind === 'move') {
      if (this.dragBase === null) this.dragBase = rn.element.style.transform || '';
      const dx = geom.x - rn.node.position.x;
      const dy = geom.y - rn.node.position.y;
      rn.element.style.transform = `translate(${dx}px, ${dy}px) ${this.dragBase}`;
    }
  }

  /** Commit a transform to the typed model and re-render the whole slide (POC granularity). */
  private commitTransform(geom: Geometry): void {
    const node = this.selectedNode;
    this.dragBase = null;
    if (!node) return;
    node.position.x = geom.x;
    node.position.y = geom.y;
    node.size.w = geom.w;
    node.size.h = geom.h;
    node.rotation = geom.rotation;
    this.render();
    this.emit();
  }

  // ---- subscription ------------------------------------------------------

  subscribe(fn: Listener): () => void {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  getState(): RenderHostState {
    return {
      loaded: !!this.presentation,
      slideCount: this.presentation?.slides.length ?? 0,
      slideIndex: this.slideIndex,
      width: this.presentation?.width ?? 0,
      height: this.presentation?.height ?? 0,
      scale: this.scale,
      selectedId: this.selectedId,
    };
  }

  private emit(): void {
    const s = this.getState();
    for (const fn of this.listeners) fn(s);
  }

  // ---- loading / rendering ----------------------------------------------

  async loadFile(data: ArrayBuffer): Promise<void> {
    const files = await parseZip(data);
    this.presentation = buildPresentation(files);
    this.slideIndex = 0;
    this.selectedId = null;
    this.render();
    this.emit();
  }

  goToSlide(index: number): void {
    if (!this.presentation) return;
    const clamped = Math.max(0, Math.min(index, this.presentation.slides.length - 1));
    if (clamped === this.slideIndex) return;
    this.slideIndex = clamped;
    this.selectedId = null;
    this.render();
    this.emit();
  }

  /** The model node currently selected, or null. */
  get selectedNode(): BaseNodeData | null {
    return this.selectedId ? (this.nodeMap.get(this.selectedId)?.node ?? null) : null;
  }

  get presentationData(): PresentationData | null {
    return this.presentation;
  }

  get currentSlide(): SlideData | null {
    return this.presentation?.slides[this.slideIndex] ?? null;
  }

  /**
   * Render the active slide, rebuilding the node map. Fits the slide to the mount width.
   * Whole-slide re-render is intentional (POC): callers invoke this after committing an edit.
   */
  render(): void {
    if (!this.presentation) return;
    const slide = this.presentation.slides[this.slideIndex];
    if (!slide) return;

    // Tear down previous render.
    this.handle?.dispose();
    this.handle = null;
    this.nodeMap.clear();
    this.topLevelIds.clear();
    for (const child of Array.from(this.stageWrap.children)) {
      if (child !== this.overlayLayer) child.remove();
    }

    this.topLevelIds = new Set(slide.nodes.map((n) => n.id));

    const handle = renderSlide(this.presentation, slide, {
      mediaUrlCache: this.mediaCache,
      onNodeRendered: (node, element) => {
        this.nodeMap.set(node.id, { node, element });
      },
    });
    this.handle = handle;

    const slideEl = handle.element;
    this.scale = this.computeFitScale();
    slideEl.style.transformOrigin = 'top left';
    slideEl.style.transform = `scale(${this.scale})`;

    // Size the wrap to the scaled slide so the overlay and scrollbars line up.
    this.stageWrap.style.width = `${this.presentation.width * this.scale}px`;
    this.stageWrap.style.height = `${this.presentation.height * this.scale}px`;
    this.overlayLayer.style.width = this.stageWrap.style.width;
    this.overlayLayer.style.height = this.stageWrap.style.height;

    this.stageWrap.insertBefore(slideEl, this.overlayLayer);
    this.syncOverlay();
  }

  private computeFitScale(): number {
    if (!this.presentation) return 1;
    const avail = this.mount.clientWidth - 32; // side gutter
    if (avail <= 0) return 1;
    return Math.min(1, avail / this.presentation.width);
  }

  // ---- selection ---------------------------------------------------------

  private handlePointerDown = (e: PointerEvent): void => {
    const target = e.target as HTMLElement | null;
    if (!target) return;
    const id = this.resolveTopLevelId(target);
    this.select(id);
    this.registerClick(id);
  };

  /** Walk up from a click target to the nearest TOP-LEVEL editable node id (groups atomic). */
  private resolveTopLevelId(from: HTMLElement): string | null {
    let cur: HTMLElement | null = from.closest('[data-node-id]');
    while (cur) {
      const id = cur.dataset.nodeId;
      if (id && this.topLevelIds.has(id)) return id;
      cur = cur.parentElement?.closest('[data-node-id]') ?? null;
    }
    return null;
  }

  select(id: string | null): void {
    if (id === this.selectedId) return;
    this.selectedId = id;
    this.syncOverlay();
    this.emit();
  }

  /** Point the transform controller at the current selection (or hide it). */
  private syncOverlay(): void {
    const node = this.selectedNode;
    this.transform.update(node ? this.nodeGeometry(node) : null);
  }

  // ---- structure edits (add / delete / reorder) --------------------------

  /** A centered box for a newly added element, in intrinsic slide px. */
  private centeredBox(w: number, h: number): { x: number; y: number; w: number; h: number } {
    const pres = this.presentation;
    const x = pres ? Math.max(0, (pres.width - w) / 2) : 0;
    const y = pres ? Math.max(0, (pres.height - h) / 2) : 0;
    return { x, y, w, h };
  }

  addTextBox(text = 'Text'): void {
    const slide = this.currentSlide;
    if (!slide) return;
    const node = createTextBox(slide, this.centeredBox(360, 90), text);
    slide.nodes.push(node);
    this.selectedId = node.id;
    this.render();
    this.emit();
  }

  addRectangle(): void {
    const slide = this.currentSlide;
    if (!slide) return;
    const node = createRectangle(slide, this.centeredBox(280, 160));
    slide.nodes.push(node);
    this.selectedId = node.id;
    this.render();
    this.emit();
  }

  deleteSelected(): void {
    const slide = this.currentSlide;
    if (!slide || !this.selectedId) return;
    if (deleteNode(slide, this.selectedId)) {
      this.selectedId = null;
      this.render();
      this.emit();
    }
  }

  reorderSelected(dir: ReorderDir): void {
    const slide = this.currentSlide;
    if (!slide || !this.selectedId) return;
    reorderNode(slide, this.selectedId, dir);
    this.render();
    this.emit();
  }

  // ---- styling (properties panel) ----------------------------------------

  /** The selected node if it is a shape (the only type the style panel edits), else null. */
  get selectedShape(): ShapeNodeData | null {
    const n = this.selectedNode;
    return n && n.nodeType === 'shape' ? (n as ShapeNodeData) : null;
  }

  applyTextStyle(patch: TextStylePatch): void {
    const shape = this.selectedShape;
    if (!shape) return;
    facadeApplyTextStyle(shape, patch);
    this.render();
    this.emit();
  }

  applyFillColor(hex: string): void {
    const shape = this.selectedShape;
    if (!shape) return;
    facadeApplyFillColor(shape, hex);
    this.render();
    this.emit();
  }

  /** Replace the selected shape's text (also used by the inline editor on commit). */
  setSelectedText(text: string): void {
    const shape = this.selectedShape;
    if (!shape) return;
    setPlainText(shape, text);
    this.render();
    this.emit();
  }

  /** Open the inline text editor for the selected shape (deterministic; no double-click). */
  editSelectedText(): void {
    const shape = this.selectedShape;
    if (shape) this.startTextEdit(shape);
  }

  // ---- inline text editing (double-click) --------------------------------

  private startTextEdit(shape: ShapeNodeData): void {
    this.cancelTextEdit();
    this.editingShape = shape;
    // Hide the transform box while editing to avoid overlap.
    this.transform.update(null);

    const ta = document.createElement('textarea');
    // Clicks inside the editor must NOT bubble to the stage (which would deselect).
    ta.addEventListener('pointerdown', (e) => e.stopPropagation());
    ta.addEventListener('mousedown', (e) => e.stopPropagation());
    ta.value = readPlainText(shape);
    const style = readTextStyle(shape);
    const k = this.scale;
    Object.assign(ta.style, {
      position: 'absolute',
      left: `${shape.position.x * k}px`,
      top: `${shape.position.y * k}px`,
      width: `${shape.size.w * k}px`,
      height: `${shape.size.h * k}px`,
      pointerEvents: 'auto',
      resize: 'none',
      boxSizing: 'border-box',
      padding: '2px',
      margin: '0',
      border: '1.5px solid #2b7fff',
      background: 'rgba(255,255,255,0.96)',
      color: style.color ? `#${style.color}` : '#000',
      fontSize: `${(style.fontSize ?? 18) * k}px`,
      fontFamily: style.fontFamily ?? 'inherit',
      lineHeight: '1.1',
      zIndex: '20',
    } satisfies Partial<CSSStyleDeclaration>);

    ta.addEventListener('keydown', (ev) => {
      ev.stopPropagation();
      // Don't treat Enter as commit while an IME composition is in progress.
      if (ev.isComposing || ev.keyCode === 229) return;
      const isEnter = ev.key === 'Enter' || ev.keyCode === 13;
      const isEsc = ev.key === 'Escape' || ev.keyCode === 27;
      if (isEnter && !ev.shiftKey) {
        ev.preventDefault();
        this.commitTextEdit();
      } else if (isEsc) {
        ev.preventDefault();
        this.cancelTextEdit();
        this.syncOverlay();
      }
    });
    // Guard against the double-click's own trailing mouseup/click blurring the freshly
    // focused textarea (which would commit-and-close immediately, looking like a no-op).
    // Ignore blurs in the first moments and just refocus; later blurs commit as intended.
    const openedAt = Date.now();
    ta.addEventListener('blur', () => {
      if (this.textEditor === ta && Date.now() - openedAt < 350) {
        setTimeout(() => {
          if (this.textEditor === ta) ta.focus();
        }, 0);
        return;
      }
      this.commitTextEdit();
    });

    this.overlayLayer.appendChild(ta);
    this.textEditor = ta;
    ta.focus();
    ta.select();
  }

  private commitTextEdit(): void {
    const ta = this.textEditor;
    const shape = this.editingShape;
    if (!ta) return;
    this.textEditor = null;
    this.editingShape = null;
    const value = ta.value;
    ta.remove();
    if (shape) {
      setPlainText(shape, value);
      // Keep the edited shape selected after commit.
      this.selectedId = shape.id;
      this.render();
      this.emit();
    }
  }

  private cancelTextEdit(): void {
    this.editingShape = null;
    if (this.textEditor) {
      this.textEditor.remove();
      this.textEditor = null;
    }
  }

  destroy(): void {
    this.cancelTextEdit();
    this.stageWrap.removeEventListener('pointerdown', this.handlePointerDown);
    document.removeEventListener('keydown', this.handleKeyDown);
    this.transform.destroy();
    this.handle?.dispose();
    this.handle = null;
    for (const url of this.mediaCache.values()) {
      try {
        URL.revokeObjectURL(url);
      } catch {
        /* ignore */
      }
    }
    this.mediaCache.clear();
    this.listeners.clear();
    this.mount.replaceChildren();
  }
}
