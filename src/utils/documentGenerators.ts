import { Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, Packer, HeadingLevel, BorderStyle } from 'docx';
import { saveAs } from 'file-saver';
import pptxgen from 'pptxgenjs';

const CENTURY_GOTHIC = 'Century Gothic';
const COLORS = {
  primary: '1B2A4A',    // Deep navy blue
  secondary: '2C3E67',  // Rich royal blue
  accent: '4A90E2',     // Bright blue
  success: '27AE60',    // Green
  warning: 'F39C12',    // Orange
  text: '2D3748',       // Dark gray
};

const parseContent = (content: string) => {
  const sections: { type: string; title?: string; content: string[] }[] = [];
  let currentSection: { type: string; title?: string; content: string[] } = {
    type: 'text',
    content: []
  };

  const lines = content.split('\n');
  
  lines.forEach(line => {
    // Detect section headers
    if (line.match(/^#{1,3}\s/)) {
      if (currentSection.content.length > 0) {
        sections.push({ ...currentSection });
      }
      const level = line.match(/^#{1,3}/)[0].length;
      currentSection = {
        type: 'heading',
        title: line.replace(/^#{1,3}\s/, ''),
        content: [],
      };
      sections.push(currentSection);
      currentSection = { type: 'text', content: [] };
    }
    // Detect lists
    else if (line.match(/^[-*]\s/)) {
      if (currentSection.type !== 'list') {
        if (currentSection.content.length > 0) {
          sections.push({ ...currentSection });
        }
        currentSection = { type: 'list', content: [] };
      }
      currentSection.content.push(line.replace(/^[-*]\s/, ''));
    }
    // Detect tables
    else if (line.includes('|')) {
      if (currentSection.type !== 'table') {
        if (currentSection.content.length > 0) {
          sections.push({ ...currentSection });
        }
        currentSection = { type: 'table', content: [] };
      }
      const cells = line.split('|').map(cell => cell.trim()).filter(cell => cell);
      if (cells.length > 0) {
        currentSection.content.push(...cells);
      }
    }
    // Regular text
    else if (line.trim()) {
      if (currentSection.type !== 'text') {
        if (currentSection.content.length > 0) {
          sections.push({ ...currentSection });
        }
        currentSection = { type: 'text', content: [] };
      }
      currentSection.content.push(line);
    }
  });

  if (currentSection.content.length > 0) {
    sections.push(currentSection);
  }

  return sections;
};

const createStyledHeading = (text: string, level: HeadingLevel): Paragraph => {
  return new Paragraph({
    text: text,
    heading: level,
    spacing: { before: 400, after: 200 },
    alignment: level === HeadingLevel.HEADING_1 ? 'center' : 'left',
    font: CENTURY_GOTHIC,
    bold: true,
    size: level === HeadingLevel.HEADING_1 ? 32 : level === HeadingLevel.HEADING_2 ? 28 : 24,
    color: COLORS.primary
  });
};

const createStyledParagraph = (text: string): Paragraph => {
  return new Paragraph({
    children: [new TextRun({
      text: text,
      font: CENTURY_GOTHIC,
      size: 24,
      color: COLORS.text
    })],
    spacing: { before: 120, after: 120, line: 360 },
    indent: { firstLine: 720 }  // 1.25 cm in twips
  });
};

const createStyledList = (items: string[]): Paragraph[] => {
  return items.map((item, index) => new Paragraph({
    children: [
      new TextRun({
        text: '• ',
        font: CENTURY_GOTHIC,
        size: 24,
        color: COLORS.accent
      }),
      new TextRun({
        text: item,
        font: CENTURY_GOTHIC,
        size: 24,
        color: COLORS.text
      })
    ],
    spacing: { before: 120, after: 120, line: 360 },
    indent: { left: 720 }
  }));
};

const createStyledTable = (data: string[]): Table => {
  const rows = [];
  const numCols = Math.ceil(Math.sqrt(data.length));
  
  for (let i = 0; i < data.length; i += numCols) {
    const rowData = data.slice(i, i + numCols);
    rows.push(new TableRow({
      children: rowData.map(cell => new TableCell({
        children: [new Paragraph({
          children: [new TextRun({
            text: cell,
            font: CENTURY_GOTHIC,
            size: 24,
            color: COLORS.text
          })],
          spacing: { line: 360 }
        })],
        margins: { top: 120, bottom: 120, left: 120, right: 120 },
        borders: {
          top: { style: BorderStyle.SINGLE, size: 1, color: COLORS.secondary },
          bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.secondary },
          left: { style: BorderStyle.SINGLE, size: 1, color: COLORS.secondary },
          right: { style: BorderStyle.SINGLE, size: 1, color: COLORS.secondary }
        }
      }))
    }));
  }

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    margins: { top: 120, bottom: 120 },
    rows
  });
};

export const generateDocx = async (content: string, filename: string) => {
  const sections = parseContent(content);
  const children: (Paragraph | Table)[] = [];

  sections.forEach(section => {
    switch (section.type) {
      case 'heading':
        children.push(createStyledHeading(
          section.title || '',
          section.title?.startsWith('# ') ? HeadingLevel.HEADING_1 :
          section.title?.startsWith('## ') ? HeadingLevel.HEADING_2 :
          HeadingLevel.HEADING_3
        ));
        break;
      case 'text':
        section.content.forEach(text => {
          children.push(createStyledParagraph(text));
        });
        break;
      case 'list':
        children.push(...createStyledList(section.content));
        break;
      case 'table':
        children.push(createStyledTable(section.content));
        break;
    }
  });

  const doc = new Document({
    sections: [{
      properties: {},
      children
    }]
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, `${filename}.docx`);
};

export const generatePptx = async (content: string, filename: string) => {
  const pres = new pptxgen();
  
  pres.layout = 'LAYOUT_WIDE';
  pres.defineLayout({ 
    name: 'LAYOUT_WIDE',
    width: 13.33,
    height: 7.5
  });

  const slides = content.split('\n\n').filter(slide => slide.trim());
  
  slides.forEach((slideContent, index) => {
    const slide = pres.addSlide();
    const [title, ...content] = slideContent.split('\n');

    slide.background = { 
      color: COLORS.primary,
      gradient: {
        type: 'linear',
        stops: [
          { color: COLORS.primary, position: 0 },
          { color: COLORS.secondary, position: 100 }
        ],
        angle: 45
      }
    };
    
    slide.addText(title.replace('Diapositiva ', ''), {
      x: 0.5,
      y: 0.5,
      w: '95%',
      h: 1.5,
      fontSize: 44,
      bold: true,
      color: 'FFFFFF',
      fontFace: CENTURY_GOTHIC,
      align: 'center'
    });

    const contentLines = content
      .map(line => line.trim())
      .filter(line => line);

    let currentY = 2.3;
    contentLines.forEach(line => {
      const isListItem = line.startsWith('- ');
      const textContent = isListItem ? line.substring(2) : line;
      
      slide.addText(textContent, {
        x: 0.5,
        y: currentY,
        w: '95%',
        h: 0.7,
        fontSize: isListItem ? 24 : 28,
        color: 'FFFFFF',
        fontFace: CENTURY_GOTHIC,
        align: isListItem ? 'left' : 'center',
        bullet: isListItem
      });

      currentY += 0.8;
    });
  });

  await pres.writeFile(`${filename}.pptx`);
};

export const generatePdf = async (content: string, filename: string) => {
  // For PDF, we'll keep it simple and just output the raw text for now
  const blob = new Blob([content], { type: 'text/plain' });
  saveAs(blob, `${filename}.pdf`);
};