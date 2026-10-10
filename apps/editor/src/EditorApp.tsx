import { useCallback, useEffect, useRef, useState } from 'react';
import { RenderHost, type RenderHostState } from './editor/RenderHost';
import { SelectionPanel } from './panels/SelectionPanel';

const INITIAL_STATE: RenderHostState = {
  loaded: false,
  slideCount: 0,
  slideIndex: 0,
  width: 0,
  height: 0,
  scale: 1,
  selectedId: null,
};

export function EditorApp() {
  const stageRef = useRef<HTMLDivElement>(null);
  const hostRef = useRef<RenderHost | null>(null);
  const [state, setState] = useState<RenderHostState>(INITIAL_STATE);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  // Create the host once the stage element exists.
  useEffect(() => {
    if (!stageRef.current) return;
    const host = new RenderHost(stageRef.current);
    hostRef.current = host;
    // Dev affordance: expose the host for console debugging / automated checks.
    (window as unknown as { __pptxHost?: RenderHost }).__pptxHost = host;
    const unsub = host.subscribe(setState);

    // Dev convenience: ?sample=<url> auto-loads a deck (e.g. a vite /@fs/ path).
    const sample = new URLSearchParams(location.search).get('sample');
    if (sample) {
      fetch(sample)
        .then((r) => r.arrayBuffer())
        .then((buf) => host.loadFile(buf))
        .catch((e) => setError(e instanceof Error ? e.message : String(e)));
    }

    return () => {
      unsub();
      host.destroy();
      hostRef.current = null;
    };
  }, []);

  const loadArrayBuffer = useCallback(async (buf: ArrayBuffer) => {
    const host = hostRef.current;
    if (!host) return;
    setBusy(true);
    setError(null);
    try {
      await host.loadFile(buf);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }, []);

  const onFile = useCallback(
    async (file: File | undefined) => {
      if (!file) return;
      loadArrayBuffer(await file.arrayBuffer());
    },
    [loadArrayBuffer],
  );

  const onDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      onFile(e.dataTransfer.files?.[0]);
    },
    [onFile],
  );

  return (
    <div className="app">
      <header className="toolbar">
        <strong className="brand">PPTX Editor</strong>
        <span className="tag">POC</span>
        <label className="btn">
          Open .pptx
          <input
            type="file"
            accept=".pptx"
            hidden
            onChange={(e) => onFile(e.target.files?.[0] ?? undefined)}
          />
        </label>
        {state.loaded && (
          <div className="nav">
            <button
              className="btn"
              disabled={state.slideIndex <= 0}
              onClick={() => hostRef.current?.goToSlide(state.slideIndex - 1)}
            >
              ‹
            </button>
            <span className="slide-count">
              {state.slideIndex + 1} / {state.slideCount}
            </span>
            <button
              className="btn"
              disabled={state.slideIndex >= state.slideCount - 1}
              onClick={() => hostRef.current?.goToSlide(state.slideIndex + 1)}
            >
              ›
            </button>
          </div>
        )}
        {state.loaded && (
          <div className="tools">
            <button className="btn" onClick={() => hostRef.current?.addTextBox()}>
              + Text
            </button>
            <button className="btn" onClick={() => hostRef.current?.addRectangle()}>
              + Rect
            </button>
            <span className="sep" />
            <button
              className="btn"
              disabled={!state.selectedId}
              onClick={() => hostRef.current?.reorderSelected('backward')}
              title="Send backward"
            >
              ⤓
            </button>
            <button
              className="btn"
              disabled={!state.selectedId}
              onClick={() => hostRef.current?.reorderSelected('forward')}
              title="Bring forward"
            >
              ⤒
            </button>
            <button
              className="btn danger"
              disabled={!state.selectedId}
              onClick={() => hostRef.current?.deleteSelected()}
              title="Delete (Del)"
            >
              🗑
            </button>
          </div>
        )}
        {busy && <span className="muted">Loading…</span>}
        {error && <span className="error">Error: {error}</span>}
      </header>

      <div className="body">
        <main
          className="stage-scroll"
          onDragOver={(e) => e.preventDefault()}
          onDrop={onDrop}
        >
          {!state.loaded && !busy && (
            <div className="empty">Drop a .pptx here, or use “Open .pptx”.</div>
          )}
          <div ref={stageRef} className="stage" />
        </main>
        <SelectionPanel host={hostRef.current} state={state} />
      </div>
    </div>
  );
}
