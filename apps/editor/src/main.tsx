import { createRoot } from 'react-dom/client';
import { EditorApp } from './EditorApp';
import './styles.css';

// Note: intentionally NOT wrapped in <StrictMode>. The editor owns imperative DOM and a
// RenderHost lifecycle in a useEffect; StrictMode's dev-only double-mount creates two hosts
// (duplicate textareas/listeners), which corrupts inline-edit commits.
createRoot(document.getElementById('root')!).render(<EditorApp />);
