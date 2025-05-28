import { Document, Paragraph, TextRun, HeadingLevel, Packer, AlignmentType, NumberingLevel, LevelFormat } from 'docx';
import { marked } from 'marked';
import { saveAs } from 'file-saver';

interface Token {
  type: string;
  text?: string;
  depth?: number;
  items?: Token[];
  tokens?: Token[];
  raw?: string;
}

export const convertMarkdownToDocx = async (markdown: string, filename: string) => {
  const tokens = marked.lexer(markdown);
  const docx = new Document({
    sections: [{
      properties: {},
      children: parseTokens(tokens)
    }]
  });

  const blob = await Packer.toBlob(docx);
  saveAs(blob, `${filename}.docx`);
};

const parseTokens = (tokens: Token[]): Paragraph[] => {
  const paragraphs: Paragraph[] = [];
  let currentListLevel = 0;

  for (const token of tokens) {
    switch (token.type) {
      case 'heading':
        paragraphs.push(new Paragraph({
          text: token.text || '',
          heading: getHeadingLevel(token.depth || 1),
          spacing: { before: 240, after: 120 },
          alignment: AlignmentType.LEFT,
          thematicBreak: token.depth === 1
        }));
        break;

      case 'paragraph':
        if (token.tokens) {
          const runs = parseInlineTokens(token.tokens);
          paragraphs.push(new Paragraph({
            children: runs,
            spacing: { before: 120, after: 120 },
            alignment: AlignmentType.JUSTIFIED
          }));
        }
        break;

      case 'list':
        currentListLevel++;
        const listItems = token.items || [];
        listItems.forEach((item, index) => {
          if (item.tokens) {
            const runs = parseInlineTokens(item.tokens);
            paragraphs.push(new Paragraph({
              children: [
                new TextRun({ text: `${getListMarker(currentListLevel, index + 1)} `, bold: true }),
                ...runs
              ],
              spacing: { before: 60, after: 60 },
              indent: { left: currentListLevel * 720 }
            }));
          }
        });
        currentListLevel--;
        break;

      case 'space':
        paragraphs.push(new Paragraph({
          text: '',
          spacing: { before: 120, after: 120 }
        }));
        break;

      case 'hr':
        paragraphs.push(new Paragraph({
          text: '',
          thematicBreak: true,
          spacing: { before: 240, after: 240 }
        }));
        break;
    }
  }

  return paragraphs;
};

const parseInlineTokens = (tokens: Token[]): TextRun[] => {
  const runs: TextRun[] = [];

  for (const token of tokens) {
    switch (token.type) {
      case 'text':
        runs.push(new TextRun({
          text: token.text || '',
          size: 24
        }));
        break;

      case 'strong':
        runs.push(new TextRun({
          text: token.text || '',
          bold: true,
          size: 24
        }));
        break;

      case 'em':
        runs.push(new TextRun({
          text: token.text || '',
          italics: true,
          size: 24
        }));
        break;

      case 'codespan':
        runs.push(new TextRun({
          text: token.text || '',
          font: 'Courier New',
          size: 24
        }));
        break;
    }
  }

  return runs;
};

const getHeadingLevel = (depth: number): HeadingLevel => {
  switch (depth) {
    case 1: return HeadingLevel.HEADING_1;
    case 2: return HeadingLevel.HEADING_2;
    case 3: return HeadingLevel.HEADING_3;
    case 4: return HeadingLevel.HEADING_4;
    case 5: return HeadingLevel.HEADING_5;
    default: return HeadingLevel.HEADING_6;
  }
};

const getListMarker = (level: number, index: number): string => {
  switch (level) {
    case 1: return `${index}.`;
    case 2: return String.fromCharCode(96 + index);
    case 3: return `${index})`;
    default: return '•';
  }
};