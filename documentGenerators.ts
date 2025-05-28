import { Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, Packer } from 'docx';
import { saveAs } from 'file-saver';
import pptxgen from 'pptxgenjs';

const parseContent = (content: string) => {
  const lines = content.split('\n');
  const elements: any[] = [];
  let currentTable: string[][] = [];
  let isInTable = false;

  lines.forEach(line => {
    if (line.includes('|')) {
      isInTable = true;
      const cells = line.split('|').map(cell => cell.trim()).filter(cell => cell);
      if (cells.length > 0) {
        currentTable.push(cells);
      }
    } else {
      if (isInTable) {
        if (currentTable.length > 0) {
          elements.push({ type: 'table', content: currentTable });
        }
        currentTable = [];
        isInTable = false;
      }

      if (line.trim()) {
        elements.push({ type: 'paragraph', content: line });
      }
    }
  });

  if (isInTable && currentTable.length > 0) {
    elements.push({ type: 'table', content: currentTable });
  }

  return elements;
};

const createFormattedText = (text: string): TextRun[] => {
  const parts = text.split(/(\*\*.*?\*\*)/g);
  return parts.map(part => {
    if (part.startsWith('**') && part.endsWith('**')) {
      return new TextRun({
        text: part.slice(2, -2),
        bold: true,
        size: 24
      });
    }
    return new TextRun({
      text: part,
      size: 24
    });
  });
};

const createTable = (tableData: string[][]): Table => {
  return new Table({
    width: {
      size: 100,
      type: WidthType.PERCENTAGE,
    },
    rows: tableData.map(row => new TableRow({
      children: row.map(cell => new TableCell({
        children: [new Paragraph({
          children: createFormattedText(cell)
        })]
      }))
    }))
  });
};

export const generateDocx = async (content: string, filename: string) => {
  const elements = parseContent(content);
  const children = elements.map(element => {
    if (element.type === 'table') {
      return createTable(element.content);
    } else {
      return new Paragraph({
        children: createFormattedText(element.content)
      });
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
  
  // Set default slide properties
  pres.layout = 'LAYOUT_WIDE';
  pres.defineLayout({ 
    name: 'LAYOUT_WIDE',
    width: 13.33,
    height: 7.5
  });

  // Enhanced color palette with dark and semi-dark colors
  const colors = {
    primary: '1B2A4A',    // Deep navy blue (base)
    secondary: '2C3E67',  // Rich royal blue
    accent1: '364B7F',    // Medium slate blue
    accent2: '1F3355',    // Dark slate blue
    accent3: '253C62',    // Steel blue
    text: 'FFFFFF'        // White
  };

  // Parse slides from content
  const slides = content.split('\n\n').filter(slide => slide.trim());
  
  slides.forEach((slideContent, index) => {
    const slide = pres.addSlide();
    const [title, ...content] = slideContent.split('\n');

    // Alternate background colors for visual interest while maintaining coherence
    const bgColor = index % 3 === 0 ? colors.primary : 
                   index % 3 === 1 ? colors.secondary : colors.accent1;

    // Set gradient background with subtle pattern
    slide.background = { 
      color: bgColor,
      gradient: {
        type: 'linear',
        stops: [
          { color: bgColor, position: 0 },
          { color: colors.accent2, position: 100 }
        ],
        angle: 45
      }
    };
    
    // Add title with Century Gothic font, centered with more space
    slide.addText(title.replace('Diapositiva ', ''), {
      x: 0.5,
      y: 0.5,
      w: '95%',
      h: 1.5,  // Increased height for better spacing
      fontSize: 44,
      bold: true,
      color: colors.text,
      fontFace: 'Century Gothic',
      align: 'center',
      glow: { size: 3, opacity: 0.3, color: colors.accent3 }
    });

    // Process content to separate subtitles and regular text
    const contentLines = content
      .map(line => line.trim())
      .filter(line => line);

    let currentY = 2.3;  // Starting position after title
    let isSubtitle = true;
    const availableHeight = 7.5 - currentY - 1.0; // Account for margins and decorative elements
    const contentHeight = contentLines.length * 0.8; // Estimate height needed
    const scaleFactor = Math.min(1, availableHeight / contentHeight);

    contentLines.forEach(line => {
      const isListItem = line.startsWith('- ');
      const textContent = isListItem ? line.substring(2) : line;
      
      // Calculate dynamic font size based on text length
      const baseSize = isSubtitle ? 26 : 20;
      const textLength = textContent.length;
      const dynamicSize = Math.min(
        baseSize,
        textLength > 100 ? baseSize * 0.8 :
        textLength > 80 ? baseSize * 0.9 : baseSize
      );

      slide.addText(textContent, {
        x: 0.5,
        y: currentY,
        w: '95%',
        h: isSubtitle ? 0.9 : 0.7,
        fontSize: Math.round(dynamicSize * scaleFactor),
        color: colors.text,
        fontFace: 'Century Gothic',
        align: isSubtitle ? 'center' : 'left',
        bullet: isListItem ? { type: 'bullet' } : false,
        bold: isSubtitle,
        glow: isSubtitle ? { size: 2, opacity: 0.2, color: colors.accent3 } : undefined,
        fit: 'shrink' // Automatically adjust text to fit within bounds
      });

      currentY += (isSubtitle ? 1.1 : 0.8) * scaleFactor;
      isSubtitle = false;
    });

    // Bottom accent bar
    slide.addShape(pres.ShapeType.rect, {
      x: 0,
      y: 6.8,
      w: '100%',
      h: 0.7,
      fill: { 
        color: colors.accent3,
        transparency: 30
      },
      line: { color: colors.accent3 }
    });

    // Top accent line
    slide.addShape(pres.ShapeType.rect, {
      x: 0,
      y: 0,
      w: '100%',
      h: 0.2,
      fill: { 
        color: colors.accent2,
        transparency: 30
      },
      line: { color: colors.accent2 }
    });

    // Subtle corner accent
    slide.addShape(pres.ShapeType.rect, {
      x: 12.5,
      y: 0,
      w: 0.8,
      h: 0.8,
      fill: { 
        color: colors.accent1,
        transparency: 40
      },
      line: { color: colors.accent1 },
      rotate: 45
    });
  });

  await pres.writeFile(`${filename}.pptx`);
};

export const generatePdf = async (content: string, filename: string) => {
  // For PDF, we'll keep it simple and just output the raw text for now
  const blob = new Blob([content], { type: 'text/plain' });
  saveAs(blob, `${filename}.pdf`);
};