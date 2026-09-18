export interface Geometry {
  x: number;
  y: number;
  w: number;
  h: number;
  rotation: number; // degrees
}

export type TransformKind = 'move' | 'resize' | 'rotate';

export interface TransformOptions {
  /** Current fit scale from intrinsic slide px → screen px. */
  getScale: () => number;
  /** Live feedback during a drag (before commit). */
  onPreview: (geom: Geometry, kind: TransformKind) => void;
  /** Final geometry when the pointer is released. */
  onCommit: (geom: Geometry, kind: TransformKind) => void;
  /**
   * Called when the box body is pressed. The box covers the selected shape, so a
   * double-click's second press lands here rather than on the shape (breaking native
   * dblclick). Return true if this press was consumed as a double-click (start editing),
   * in which case the move is aborted.
   */
  onBoxPointerDown?: () => boolean;
}

const MIN_SIZE = 8; // intrinsic px

// 8 resize handles as unit vectors from center: x/y ∈ {-1,0,1}, excluding (0,0).
const HANDLES: Array<{ hx: -1 | 0 | 1; hy: -1 | 0 | 1; cursor: string }> = [
  { hx: -1, hy: -1, cursor: 'nwse-resize' },
  { hx: 0, hy: -1, cursor: 'ns-resize' },
  { hx: 1, hy: -1, cursor: 'nesw-resize' },
  { hx: 1, hy: 0, cursor: 'ew-resize' },
  { hx: 1, hy: 1, cursor: 'nwse-resize' },
  { hx: 0, hy: 1, cursor: 'ns-resize' },
  { hx: -1, hy: 1, cursor: 'nesw-resize' },
  { hx: -1, hy: 0, cursor: 'ew-resize' },
];

/**
 * Draws the selection box, 8 resize handles and a rotate handle over a node, and turns
 * pointer drags into new {x,y,w,h,rotation}. Math is rotation-aware: resize keeps the
 * opposite corner/edge fixed even when the node is rotated. All geometry is in intrinsic
 * slide px; the controller converts pointer deltas through the current scale.
 */
export class TransformController {
  private box: HTMLElement;
  private geom: Geometry | null = null;

  constructor(
    private layer: HTMLElement,
    private opts: TransformOptions,
  ) {
    this.box = document.createElement('div');
    this.box.style.position = 'absolute';
    this.box.style.boxSizing = 'border-box';
    this.box.style.border = '1.5px solid #2b7fff';
    this.box.style.outline = '1px solid rgba(255,255,255,0.55)';
    this.box.style.pointerEvents = 'auto';
    this.box.style.cursor = 'move';
    this.box.style.display = 'none';
    this.box.style.touchAction = 'none';
    this.box.addEventListener('pointerdown', (e) => this.startMove(e));

    for (const h of HANDLES) {
      // Generous transparent hit area (18px) with a smaller visible dot centered inside.
      const el = document.createElement('div');
      el.dataset.hx = String(h.hx);
      el.dataset.hy = String(h.hy);
      el.style.position = 'absolute';
      el.style.width = '18px';
      el.style.height = '18px';
      el.style.marginLeft = '-9px';
      el.style.marginTop = '-9px';
      el.style.left = `${(h.hx + 1) * 50}%`;
      el.style.top = `${(h.hy + 1) * 50}%`;
      el.style.display = 'flex';
      el.style.alignItems = 'center';
      el.style.justifyContent = 'center';
      el.style.pointerEvents = 'auto';
      el.style.touchAction = 'none';
      el.style.cursor = h.cursor;
      const dot = document.createElement('div');
      dot.style.width = '10px';
      dot.style.height = '10px';
      dot.style.background = '#fff';
      dot.style.border = '1.5px solid #2b7fff';
      dot.style.borderRadius = '2px';
      dot.style.pointerEvents = 'none';
      el.appendChild(dot);
      el.addEventListener('pointerdown', (e) => this.startResize(e, h.hx, h.hy));
      this.box.appendChild(el);
    }

    const rot = document.createElement('div');
    rot.style.position = 'absolute';
    rot.style.width = '12px';
    rot.style.height = '12px';
    rot.style.marginLeft = '-6px';
    rot.style.left = '50%';
    rot.style.top = '-26px';
    rot.style.background = '#fff';
    rot.style.border = '1.5px solid #2b7fff';
    rot.style.borderRadius = '50%';
    rot.style.pointerEvents = 'auto';
    rot.style.touchAction = 'none';
    rot.style.cursor = 'grab';
    rot.addEventListener('pointerdown', (e) => this.startRotate(e));
    this.box.appendChild(rot);

    this.layer.appendChild(this.box);
  }

  /** Update the box to reflect a node's geometry (or hide it when null). */
  update(geom: Geometry | null): void {
    this.geom = geom;
    if (!geom) {
      this.box.style.display = 'none';
      return;
    }
    this.paint(geom);
  }

  private paint(g: Geometry): void {
    const k = this.opts.getScale();
    this.box.style.display = 'block';
    this.box.style.left = `${g.x * k}px`;
    this.box.style.top = `${g.y * k}px`;
    this.box.style.width = `${g.w * k}px`;
    this.box.style.height = `${g.h * k}px`;
    this.box.style.transform = g.rotation ? `rotate(${g.rotation}deg)` : '';
    this.box.style.transformOrigin = 'center center';
  }

  // ---- drags -------------------------------------------------------------

  private startMove(e: PointerEvent): void {
    if (!this.geom) return;
    // Only start a move when the box body itself is grabbed (handles stop propagation).
    if (e.target !== this.box) return;
    e.preventDefault();
    e.stopPropagation();
    // A double-click's second press lands on the box; let the host consume it as an edit.
    if (this.opts.onBoxPointerDown?.()) return;
    const start = this.geom;
    const sx = e.clientX;
    const sy = e.clientY;
    this.drag(e, 'move', (cx, cy) => {
      const k = this.opts.getScale();
      return { ...start, x: start.x + (cx - sx) / k, y: start.y + (cy - sy) / k };
    });
  }

  private startResize(e: PointerEvent, hx: number, hy: number): void {
    if (!this.geom) return;
    e.preventDefault();
    e.stopPropagation();
    const start = this.geom;
    const sx = e.clientX;
    const sy = e.clientY;
    const r = (start.rotation * Math.PI) / 180;
    const cos = Math.cos(r);
    const sin = Math.sin(r);
    const cx0 = start.x + start.w / 2;
    const cy0 = start.y + start.h / 2;
    this.drag(e, 'resize', (cx, cy) => {
      const k = this.opts.getScale();
      const dx = (cx - sx) / k;
      const dy = (cy - sy) / k;
      // Rotate the world delta into the node's local (unrotated) frame: R(-r)·d.
      const ldx = dx * cos + dy * sin;
      const ldy = -dx * sin + dy * cos;
      const newW = Math.max(MIN_SIZE, start.w + hx * ldx);
      const newH = Math.max(MIN_SIZE, start.h + hy * ldy);
      // Shift the center so the opposite corner/edge stays fixed (in world space).
      const shiftLx = (hx * (newW - start.w)) / 2;
      const shiftLy = (hy * (newH - start.h)) / 2;
      const cx1 = cx0 + shiftLx * cos - shiftLy * sin;
      const cy1 = cy0 + shiftLx * sin + shiftLy * cos;
      return { ...start, w: newW, h: newH, x: cx1 - newW / 2, y: cy1 - newH / 2 };
    });
  }

  private startRotate(e: PointerEvent): void {
    if (!this.geom) return;
    e.preventDefault();
    e.stopPropagation();
    const start = this.geom;
    const rect = this.layer.getBoundingClientRect();
    const k = this.opts.getScale();
    const ccx = rect.left + (start.x + start.w / 2) * k;
    const ccy = rect.top + (start.y + start.h / 2) * k;
    const startAngle = Math.atan2(e.clientY - ccy, e.clientX - ccx);
    this.drag(e, 'rotate', (cx, cy) => {
      const ang = Math.atan2(cy - ccy, cx - ccx);
      let deg = start.rotation + ((ang - startAngle) * 180) / Math.PI;
      // Snap to 15° increments while Shift is held would go here; keep it simple for POC.
      deg = ((deg % 360) + 360) % 360;
      return { ...start, rotation: deg };
    });
  }

  private drag(
    e: PointerEvent,
    kind: TransformKind,
    compute: (clientX: number, clientY: number) => Geometry,
  ): void {
    const target = e.currentTarget as HTMLElement;
    target.setPointerCapture?.(e.pointerId);
    const onMove = (ev: PointerEvent) => {
      const g = compute(ev.clientX, ev.clientY);
      this.geom = g;
      this.paint(g);
      this.opts.onPreview(g, kind);
    };
    const onUp = (ev: PointerEvent) => {
      target.releasePointerCapture?.(ev.pointerId);
      target.removeEventListener('pointermove', onMove);
      target.removeEventListener('pointerup', onUp);
      if (this.geom) this.opts.onCommit(this.geom, kind);
    };
    target.addEventListener('pointermove', onMove);
    target.addEventListener('pointerup', onUp);
  }

  destroy(): void {
    this.box.remove();
  }
}
