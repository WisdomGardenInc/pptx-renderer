/**
 * End-to-end: parse a real .pptx, edit the model, write the slide back, and
 * reassemble the package. The assertion that matters is not made here — the
 * artifact is handed to PowerPoint's own parser (python-pptx) by the companion
 * script — but everything up to the bytes is covered.
 */
import { describe, expect, it } from 'vitest';
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import JSZip from 'jszip';
import { parseZip } from '../../../src/parser/ZipParser';
import { buildPresentation } from '../../../src/model/Presentation';
import { materializeSlideData } from '../../../src/model/Slide';
import type { ShapeNodeData } from '../../../src/model/nodes/ShapeNode';
import { serializeSlide } from '../../../src/writer/SlideWriter';
import { emuToPx } from '../../../src/parser/units';
import { SafeXmlNode } from '../../../src/parser/XmlParser';

const SRC = resolve(__dirname, '../../../docs/example/1-chart-and-complex/source.pptx');
const OUT = process.env.WRITER_OUT ?? resolve(__dirname, '../../../docs/agent-tmp/writer-out.pptx');
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

describe('writer round-trip on a real deck', () => {
  it('edits a slide and reassembles a package', async () => {
    // Copy into a plain ArrayBuffer: jsdom and Node hold separate realms, and
    // JSZip rejects a Buffer whose backing store came from the other one.
    const buf = readFileSync(SRC);
    const bytes = new Uint8Array(buf.length);
    bytes.set(buf);
    const files = await parseZip(bytes.buffer);
    const pres = buildPresentation(files);

    const slide = pres.slides[0];
    materializeSlideData(slide);
    expect(slide.nodes.length).toBeGreaterThan(0);

    // Find a shape with text — the edit has to be observable from outside.
    const target = slide.nodes.find(
      (n): n is ShapeNodeData =>
        n.nodeType === 'shape' && !!(n as ShapeNodeData).textBody?.paragraphs.length,
    );
    expect(target).toBeDefined();

    target!.position.x = emuToPx(914400);
    target!.position.y = emuToPx(457200);
    const para = target!.textBody!.paragraphs[0];
    para.runs = [{ text: 'EDITED-BY-WRITER', properties: para.runs[0]?.properties }];

    const doc = target!.source.element!.ownerDocument;
    const sf = doc.createElementNS(A, 'a:solidFill');
    const clr = doc.createElementNS(A, 'a:srgbClr');
    clr.setAttribute('val', '00B050');
    sf.appendChild(clr);
    target!.fill = new SafeXmlNode(sf);

    const { path, xml } = serializeSlide(slide);
    expect(path).toBe('ppt/slides/slide1.xml');
    expect(xml).toContain('EDITED-BY-WRITER');
    expect(xml).toContain('<a:off x="914400" y="457200"/>');
    expect(xml).toContain('val="00B050"');

    // Reassemble: baseline + this one overridden part. Everything else is copied.
    const zip = await JSZip.loadAsync(bytes);
    // createFolders:false — JSZip otherwise adds directory entries ("ppt/",
    // "ppt/slides/") that the original package does not carry. OPC has no use for
    // them, and a package should not grow entries just from being edited.
    zip.file(path, xml, { createFolders: false });
    const out = await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });

    mkdirSync(dirname(OUT), { recursive: true });
    writeFileSync(OUT, out);
    expect(out.length).toBeGreaterThan(1000);
  });
});
