// Library
import { marked } from 'marked';
import PptxGenJS from 'pptxgenjs';

// Styles
import "./style.css"

const markdown = /** @type HTMLTextAreaElement */ (document.getElementById('markdown'))!;
const status = document.getElementById('status')!;

document.getElementById('generate')?.addEventListener('click', () => {
  //@ts-ignore for now
  generatePPT(markdown.value);
});

async function generatePPT(markdownContent: string) {
  const pptx = new PptxGenJS();
  const slidesContent = markdownContent.split('---');

  slidesContent.forEach(async content => {
    const slide = pptx.addSlide();
    const parsedMarkdown = await marked(content.trim());
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = parsedMarkdown;

    let y = 1; // Y position for text elements

    Array.from(tempDiv.childNodes).forEach(node => {
      if (node.nodeName === 'H1') {
        slide.addText(node.textContent ?? "", { x: 1, y: y, fontSize: 24 });
        y += 1;
      } else if (node.nodeName === 'H2') {
        slide.addText(node.textContent ?? "", { x: 1, y: y, fontSize: 20 });
        y += 0.5;
      } else if (node.nodeName === 'P') {
        slide.addText(node.textContent ?? "", { x: 1, y: y, fontSize: 16 });
        y += 0.5;
      } else if (node.nodeName === 'UL') {
        node.childNodes.forEach(li => {
          slide.addText(`• ${li.textContent}`, { x: 1, y: y, fontSize: 16 });
          y += 0.5;
        });
      }
    });
  });

  pptx.writeFile({ fileName: 'presentation.pptx' }).then(() => {
    status.textContent = 'Presentation generated!';
  }).catch(err => {
    status.textContent = 'Error generating presentation.';
    console.error(err);
  });
}
