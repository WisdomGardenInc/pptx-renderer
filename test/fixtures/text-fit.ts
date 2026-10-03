import type { PptxFiles } from '../../src/parser/ZipParser';
import { groupPlaceholderFiles, transform } from './group-placeholder-transform';

const LONG =
  'A replacement sentence much longer than the sample text this box was originally sized for';

function textBox(id: number, y: number, autofit: string, text: string): string {
  return `<p:sp>
    <p:nvSpPr><p:cNvPr id="${id}" name="Box ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr><a:xfrm>${transform(40, y, 220, 30)}</a:xfrm><a:prstGeom prst="rect"/></p:spPr>
    <p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0">${autofit}</a:bodyPr><a:lstStyle/>
      <a:p><a:r><a:rPr sz="1400"/><a:t>${text}</a:t></a:r></a:p></p:txBody>
  </p:sp>`;
}

/** One slide: text that fits, text the renderer shrinks, text that spills out. */
export function textFitFiles(): PptxFiles {
  return groupPlaceholderFiles(
    textBox(2, 20, '<a:spAutoFit/>', 'Short') +
      textBox(3, 120, '<a:normAutofit/>', LONG) +
      textBox(4, 220, '<a:noAutofit/>', LONG),
  );
}
