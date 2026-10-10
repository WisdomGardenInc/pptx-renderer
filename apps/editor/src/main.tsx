import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { EditorApp } from './EditorApp';
import './styles.css';

// StrictMode is on deliberately: its dev-only double-mount (effect → cleanup → effect)
// is the cheapest standing check that RenderHost's teardown is idempotent — and hosts
// that embed this editor, Next.js among them, enable it by default anyway.
createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <EditorApp />
  </StrictMode>,
);
