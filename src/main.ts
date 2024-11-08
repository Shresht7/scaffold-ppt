// Library
import { marked } from 'marked';
import PptxGenJS from 'pptxgenjs';

// Styles
import "./style.css"

const btnGenerate = document.getElementById('generate') as HTMLButtonElement | null;
const markdown = document.getElementById('markdown') as HTMLTextAreaElement | null;
const status = document.getElementById('status') as HTMLElement | null;

if (markdown && status) {
  btnGenerate?.addEventListener('click', () => {
    generatePPT(markdown.value);
  });
}

markdown?.addEventListener('input', () => {
  btnGenerate!.disabled = !(markdown.value !== "");
  console.log(btnGenerate!.disabled)
})

async function generatePPT(markdownContent: string) {
  const pptx = new PptxGenJS();
  const slidesContent = markdownContent.split('---');

  // Use for...of loop to properly await async operations
  for (const content of slidesContent) {
    const slide = pptx.addSlide();
    const parsedMarkdown = await marked(content.trim());
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = parsedMarkdown;

    let y = 1; // Y position for text elements

    // Iterate through the child nodes of the parsed markdown
    for (const node of tempDiv.childNodes) {
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
        // Handle unordered lists
        for (const li of node.childNodes) {
          if (li.nodeName === 'LI') {
            slide.addText(`• ${li.textContent ?? ""}`, { x: 1, y: y, fontSize: 16 });
            y += 0.5;
          }
        }
      }
    }
  }

  // Generate the PowerPoint file
  try {
    await pptx.writeFile({ fileName: 'presentation.pptx' });
    if (status) status.textContent = 'Presentation generated!';
  } catch (err) {
    if (status) status.textContent = 'Error generating presentation.';
    console.error(err);
  }
}
