import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { composeDeck } from '../../../src/compose/composeDeck';
import type { ComposeOp } from '../../../src/compose/types';
import type { GroupNodeData } from '../../../src/model/nodes/GroupNode';
import type { PicNodeData } from '../../../src/model/nodes/PicNode';
import type { SlideData } from '../../../src/model/Slide';
import { resolveRelTarget } from '../../../src/parser/RelParser';
import {
  groupBox,
  open,
  picture,
  slideXml,
  templateWithPicture,
  templateWithSlide,
  textShape,
} from '../../fixtures/compose-template';

const PNG_HEAD = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
const PNG = new Uint8Array([...PNG_HEAD, 1, 2, 3]);
const OTHER_PNG = new Uint8Array([...PNG_HEAD, 9, 9]);
const BOX: [number, number, number, number] = [64, 160, 480, 300];
const PLAIN = picture({ id: 3, name: 'Art', box: BOX, embed: 'rId3' });
const TWO_ART = slideXml(
  textShape(2, 'Title', [64, 64, 800, 60], 'Two frames') +
    picture({ id: 3, name: 'Left', box: [64, 160, 320, 240], embed: 'rId3' }) +
    picture({ id: 4, name: 'Right', box: [420, 160, 320, 240], embed: 'rId3' }),
);
const POSTERS: Array<{ label: string; tag: string; kind: string }> = [
  { label: 'a video poster', tag: 'videoFile', kind: 'video' },
  { label: 'a quick-time poster', tag: 'quickTimeFile', kind: 'video' },
  { label: 'an audio poster', tag: 'audioFile', kind: 'audio' },
  { label: 'a wav audio poster', tag: 'wavAudioFile', kind: 'audio' },
  { label: 'an audio-cd poster', tag: 'audioCd', kind: 'audio' },
];

function copy(bytes: Uint8Array): ArrayBuffer {
  const out = new Uint8Array(bytes.length);
  out.set(bytes);
  return out.buffer;
}

function pictureOf(slide: SlideData): PicNodeData {
  const node = slide.nodes.find((candidate) => candidate.nodeType === 'picture');
  if (!node) throw new Error('the composed slide has no picture');
  return node as PicNodeData;
}

/** The pictures of a slide, in the order the slide lists them. */
function picturesOf(slide: SlideData): PicNodeData[] {
  return slide.nodes.filter((node) => node.nodeType === 'picture') as PicNodeData[];
}

/** The blip a picture inside a group points at, read through the group's own children. */
function groupPictureEmbed(slide: SlideData): string {
  const group = slide.nodes.find((node) => node.nodeType === 'group') as GroupNodeData | undefined;
  if (!group) throw new Error('the composed slide has no group');
  for (const child of group.children) {
    const embed = child.child('blipFill').child('blip').attr('r:embed');
    if (embed) return embed;
  }
  throw new Error('the group has no picture');
}

/** The package part a relationship of `slide` resolves to. */
function partFor(slide: SlideData, relId: string): string {
  const rel = slide.rels.get(relId);
  if (!rel) throw new Error(`the slide has no relationship ${relId}`);
  return resolveRelTarget('ppt/slides', rel.target);
}

async function partBytes(deck: Uint8Array, path: string): Promise<Uint8Array> {
  const file = (await JSZip.loadAsync(copy(deck))).file(path);
  if (!file) throw new Error(`the deck has no part ${path}`);
  return file.async('uint8array');
}

async function partText(deck: Uint8Array, path: string): Promise<string> {
  const file = (await JSZip.loadAsync(copy(deck))).file(path);
  if (!file) throw new Error(`the deck has no part ${path}`);
  return file.async('string');
}

/** The media parts a package carries, by name. */
async function mediaParts(deck: Uint8Array): Promise<string[]> {
  const zip = await JSZip.loadAsync(copy(deck));
  return Object.keys(zip.files).filter((name) => name.startsWith('ppt/media/'));
}

async function compose(pic: string, ops: ComposeOp[]) {
  return composeDeck(await templateWithPicture(pic), [{ source: 1, ops }]);
}

describe('set_image', () => {
  it('shows the new image and leaves the media the template had alone', async () => {
    const source = await templateWithPicture(PLAIN);
    const original = await partBytes(source, 'ppt/media/image1.png');

    const result = await compose(PLAIN, [{ op: 'set_image', element: '3', bytes: PNG }]);

    expect(result.problems).toEqual([]);
    const pres = await open(result.bytes);
    const pic = pictureOf(pres.slides[0]);
    const part = partFor(pres.slides[0], pic.blipEmbed!);
    expect(part).toMatch(/^ppt\/media\/picture\d+\.png$/);
    expect(await partBytes(result.bytes, part)).toEqual(PNG);
    expect(await partBytes(result.bytes, 'ppt/media/image1.png')).toEqual(original);
  });

  it('drops the crop and the vector alternative the picture carried', async () => {
    const cropped = picture({
      id: 3,
      name: 'Art',
      box: BOX,
      embed: 'rId3',
      crop: 'l="10000" t="5000"',
      svgEmbed: 'rId4',
    });
    const before = pictureOf((await open(await templateWithPicture(cropped))).slides[0]);
    expect(before.blipEmbed).toBe('rId4');
    expect(before.crop).toEqual({ top: 0.05, bottom: 0, left: 0.1, right: 0 });
    const media = await mediaParts(await templateWithPicture(cropped));

    const result = await compose(cropped, [{ op: 'set_image', element: '3', bytes: PNG }]);

    expect(result.problems).toEqual([]);
    const pres = await open(result.bytes);
    const pic = pictureOf(pres.slides[0]);
    const part = partFor(pres.slides[0], pic.blipEmbed!);
    expect(pic.crop).toBeUndefined();
    expect(media).not.toContain(part);
    expect(await partBytes(result.bytes, part)).toEqual(PNG);
    expect(await partText(result.bytes, 'ppt/slides/slide1.xml')).not.toContain('svgBlip');
  });

  it('replaces the image inside a group', async () => {
    const result = await compose(groupBox(4, 'Icons', BOX, PLAIN), [
      { op: 'set_image', element: '3', bytes: PNG },
    ]);

    expect(result.problems).toEqual([]);
    const pres = await open(result.bytes);
    const part = partFor(pres.slides[0], groupPictureEmbed(pres.slides[0]));
    expect(part).toMatch(/^ppt\/media\/picture\d+\.png$/);
    expect(await partBytes(result.bytes, part)).toEqual(PNG);
  });

  it('gives each copy of a repeated slide its own image', async () => {
    const result = await composeDeck(await templateWithPicture(PLAIN), [
      { source: 1, ops: [{ op: 'set_image', element: '3', bytes: PNG }] },
      { source: 1, ops: [{ op: 'set_image', element: '3', bytes: OTHER_PNG }] },
    ]);

    expect(result.problems).toEqual([]);
    const pres = await open(result.bytes);
    expect(pres.slides).toHaveLength(2);
    const first = partFor(pres.slides[0], pictureOf(pres.slides[0]).blipEmbed!);
    const second = partFor(pres.slides[1], pictureOf(pres.slides[1]).blipEmbed!);
    expect(first).not.toBe(second);
    expect(await partBytes(result.bytes, first)).toEqual(PNG);
    expect(await partBytes(result.bytes, second)).toEqual(OTHER_PNG);
  });

  it('replaces two pictures of one page without confusing them', async () => {
    const result = await composeDeck(await templateWithSlide(TWO_ART), [
      {
        source: 1,
        ops: [
          { op: 'set_image', element: '3', bytes: PNG },
          { op: 'set_image', element: '4', bytes: OTHER_PNG },
        ],
      },
    ]);

    expect(result.problems).toEqual([]);
    const pres = await open(result.bytes);
    const pics = picturesOf(pres.slides[0]);
    expect(pics.map((pic) => pic.id)).toEqual(['3', '4']);
    const parts = pics.map((pic) => partFor(pres.slides[0], pic.blipEmbed!));
    expect(parts).toHaveLength(2);
    expect(parts[0]).not.toBe(parts[1]);
    expect(await partBytes(result.bytes, parts[0])).toEqual(PNG);
    expect(await partBytes(result.bytes, parts[1])).toEqual(OTHER_PNG);
  });

  it('reports a picture an earlier operation took off the page', async () => {
    const deck = await templateWithSlide(TWO_ART);
    const media = await mediaParts(deck);

    const result = await composeDeck(deck, [
      {
        source: 1,
        ops: [
          {
            op: 'set_items',
            group: { direction: 'x', members: [{ art: '3' }, { art: '4' }] },
            items: [{}],
          },
          { op: 'set_image', element: '4', bytes: PNG },
        ],
      },
    ]);

    expect(result.problems.map((problem) => problem.message)).toEqual([
      'element 4 is no longer on the page',
    ]);
    const pres = await open(result.bytes);
    expect(pres.slides[0].nodes.some((node) => node.id === '4')).toBe(false);
    expect((await mediaParts(result.bytes)).sort()).toEqual([...media].sort());
  });

  it('takes over a vector-only picture and drops its vector alternative', async () => {
    const vector = picture({ id: 3, name: 'Art', box: BOX, svgEmbed: 'rId3' });
    const before = pictureOf((await open(await templateWithPicture(vector))).slides[0]);
    expect(before.blipEmbed).toBe('rId3');

    const result = await compose(vector, [{ op: 'set_image', element: '3', bytes: PNG }]);

    expect(result.problems).toEqual([]);
    const pres = await open(result.bytes);
    const pic = pictureOf(pres.slides[0]);
    const part = partFor(pres.slides[0], pic.blipEmbed!);
    expect(await partBytes(result.bytes, part)).toEqual(PNG);
    expect(await partText(result.bytes, 'ppt/slides/slide1.xml')).not.toContain('svgBlip');
  });

  it('takes over a picture whose blip only links to its image', async () => {
    const linked = picture({ id: 3, name: 'Art', box: BOX, link: 'rId3' });
    const before = pictureOf((await open(await templateWithPicture(linked))).slides[0]);
    expect(before.blipLink).toBe('rId3');
    expect(before.blipEmbed).toBeUndefined();

    const result = await compose(linked, [{ op: 'set_image', element: '3', bytes: PNG }]);

    expect(result.problems).toEqual([]);
    const pres = await open(result.bytes);
    const pic = pictureOf(pres.slides[0]);
    expect(pic.blipLink).toBeUndefined();
    expect(await partBytes(result.bytes, partFor(pres.slides[0], pic.blipEmbed!))).toEqual(PNG);
    expect(await partText(result.bytes, 'ppt/slides/slide1.xml')).not.toContain('r:link');
  });

  const REFUSALS: Array<{ label: string; pic: string; op: ComposeOp; message: string }> = [
    {
      label: 'a picture that is not on the slide',
      pic: PLAIN,
      op: { op: 'set_image', element: '99', bytes: PNG },
      message: 'element 99 is not an element of the slide',
    },
    {
      label: 'a text shape',
      pic: PLAIN,
      op: { op: 'set_image', element: '2', bytes: PNG },
      message: 'element 2 is a shape, not a picture',
    },
    ...POSTERS.map(
      ({ label, tag, kind }): { label: string; pic: string; op: ComposeOp; message: string } => ({
        label,
        pic: picture({ id: 3, name: 'Art', box: BOX, embed: 'rId3', poster: { tag, rel: 'rId4' } }),
        op: { op: 'set_image', element: '3', bytes: PNG },
        message: `element 3 is a ${kind} poster, not a picture`,
      }),
    ),
    {
      label: 'bytes that are not an image',
      pic: PLAIN,
      op: { op: 'set_image', element: '3', bytes: new Uint8Array([1, 2, 3]) },
      message:
        'element 3 could not take the new image: Unsupported image bytes: expected PNG or JPEG',
    },
  ];

  it.each(REFUSALS)('reports $label and still produces the deck', async ({ pic, op, message }) => {
    const result = await compose(pic, [op]);

    expect(result.problems.map((problem) => problem.message)).toEqual([message]);
    expect(pictureOf((await open(result.bytes)).slides[0]).blipEmbed).toBe('rId3');
  });
});
