import type { RenderHost, RenderHostState } from '../editor/RenderHost';
import { readFillColor, readTextStyle } from '../editor/StyleFacade';

interface Props {
  host: RenderHost | null;
  state: RenderHostState;
}

const FONTS = ['Arial', 'Calibri', 'Times New Roman', 'Helvetica', 'Georgia', 'Courier New'];

/** Right-hand inspector + properties editor for the selected node. */
export function SelectionPanel({ host, state }: Props) {
  const node = host?.selectedNode ?? null;
  const shape = host?.selectedShape ?? null;

  return (
    <aside className="panel">
      <h2 className="panel-title">Inspector</h2>
      {!state.loaded && <p className="muted">No presentation loaded.</p>}
      {state.loaded && !node && <p className="muted">Click an element to select it.</p>}

      {node && (
        <>
          <dl className="props">
            <Row label="Name" value={node.name || '(unnamed)'} />
            <Row label="Type" value={node.nodeType} />
            <Row label="X" value={round(node.position.x)} />
            <Row label="Y" value={round(node.position.y)} />
            <Row label="W" value={round(node.size.w)} />
            <Row label="H" value={round(node.size.h)} />
            <Row label="Rotation" value={`${round(node.rotation)}°`} />
          </dl>

          {shape && <ShapeStyleControls host={host!} shape={shape} />}
        </>
      )}
    </aside>
  );
}

function ShapeStyleControls({
  host,
  shape,
}: {
  host: RenderHost;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  shape: any;
}) {
  const text = readTextStyle(shape);
  const fill = readFillColor(shape);

  return (
    <div className="style">
      <button className="btn edit-text-btn" onClick={() => host.editSelectedText()}>
        ✏️ Edit text
      </button>

      <h3 className="section">Fill</h3>
      <div className="field">
        <label>Color</label>
        <input
          type="color"
          value={`#${fill ?? 'FFFFFF'}`}
          onChange={(e) => host.applyFillColor(e.target.value)}
        />
      </div>

      <h3 className="section">Text</h3>
      <div className="field">
        <label>Color</label>
        <input
          type="color"
          value={`#${text.color ?? '000000'}`}
          onChange={(e) => host.applyTextStyle({ color: e.target.value })}
        />
      </div>
      <div className="field">
        <label>Size</label>
        <input
          type="number"
          min={4}
          max={400}
          value={text.fontSize ?? 18}
          onChange={(e) => host.applyTextStyle({ fontSize: Number(e.target.value) })}
        />
      </div>
      <div className="field">
        <label>Font</label>
        <select
          value={FONTS.includes(text.fontFamily ?? '') ? text.fontFamily : ''}
          onChange={(e) => e.target.value && host.applyTextStyle({ fontFamily: e.target.value })}
        >
          <option value="">{text.fontFamily ?? '(inherit)'}</option>
          {FONTS.map((f) => (
            <option key={f} value={f}>
              {f}
            </option>
          ))}
        </select>
      </div>
      <div className="field">
        <label>Style</label>
        <div className="toggles">
          <button
            className={`btn toggle ${text.bold ? 'on' : ''}`}
            onClick={() => host.applyTextStyle({ bold: !text.bold })}
            style={{ fontWeight: 700 }}
          >
            B
          </button>
          <button
            className={`btn toggle ${text.italic ? 'on' : ''}`}
            onClick={() => host.applyTextStyle({ italic: !text.italic })}
            style={{ fontStyle: 'italic' }}
          >
            I
          </button>
        </div>
      </div>
      <p className="hint">
        Or double-click the element to edit text. Press Enter to apply, Esc to cancel.
      </p>
    </div>
  );
}

function Row({ label, value }: { label: string; value: string | number }) {
  return (
    <>
      <dt>{label}</dt>
      <dd>{value}</dd>
    </>
  );
}

function round(n: number): number {
  return Math.round(n * 10) / 10;
}
