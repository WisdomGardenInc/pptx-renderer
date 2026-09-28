/**
 * The editor core, as consumed by embedding hosts.
 *
 * Everything here is imperative and framework-free on purpose: `RenderHost` owns real
 * DOM and a slide render handle, which no component tree should be re-creating on its
 * own schedule. A host wraps it in whatever it uses — see `apps/editor/src/EditorApp.tsx`
 * for the React shell — and is responsible for calling `destroy()` exactly once.
 */

export { RenderHost } from './RenderHost';
export type { RenderHostState } from './RenderHost';
export { TransformController } from './TransformController';
export type { Geometry, TransformKind } from './TransformController';
export {
  applyTextStyle,
  applyFillColor,
  setPlainText,
  readPlainText,
  readTextStyle,
} from './StyleFacade';
export type { TextStyle, TextStylePatch } from './StyleFacade';
export { createTextBox, createRectangle, deleteNode, reorderNode, freshId } from './nodeOps';
export type { ReorderDir } from './nodeOps';
