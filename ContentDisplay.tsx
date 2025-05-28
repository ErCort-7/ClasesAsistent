import React, { useState } from 'react';
import { Download, Copy, CheckCircle2, FileText, Presentation, FileQuestion } from 'lucide-react';
import { generateDocx, generatePptx, generatePdf } from '../../utils/documentGenerators';

interface ContentDisplayProps {
  title: string;
  content: string | null;
  type: 'guion' | 'presentacion' | 'ejercicios';
  isLoading: boolean;
  format: string;
}

const ContentDisplay: React.FC<ContentDisplayProps> = ({ title, content, type, isLoading, format }) => {
  const [copied, setCopied] = React.useState(false);
  const [showPreview, setShowPreview] = useState(false);

  const handleCopy = () => {
    if (content) {
      navigator.clipboard.writeText(content);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }
  };

  const handleDownload = async () => {
    if (!content) return;
    
    const filename = title.toLowerCase().replace(/\s+/g, '-');
    
    try {
      switch (format) {
        case 'docx':
          await generateDocx(content, filename);
          break;
        case 'pptx':
          await generatePptx(content, filename);
          break;
        case 'pdf':
          await generatePdf(content, filename);
          break;
        default:
          const blob = new Blob([content], { type: 'text/plain' });
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = `${filename}.txt`;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          URL.revokeObjectURL(url);
      }
    } catch (error) {
      console.error('Error generating document:', error);
      alert('Error generating document. Please try again.');
    }
  };

  const renderPresentationPreview = () => {
    if (!content) return null;

    const slides = content.split('\n\n').filter(slide => slide.trim());
    
    return (
      <div className="space-y-4">
        {slides.map((slide, index) => {
          const [title, ...content] = slide.split('\n');
          return (
            <div key={index} className="border rounded-lg p-4 bg-white shadow-sm">
              <h3 className="font-bold text-lg mb-2">{title.replace('Diapositiva ', '')}</h3>
              <div className="space-y-2">
                {content.map((line, i) => (
                  <p key={i} className="text-sm">
                    {line.startsWith('- ') ? line.substring(2) : line}
                  </p>
                ))}
              </div>
            </div>
          );
        })}
      </div>
    );
  };

  let icon;
  let bgColor;
  
  switch (type) {
    case 'guion':
      icon = <FileText className="h-6 w-6 text-blue-600" />;
      bgColor = 'bg-blue-50 border-blue-200';
      break;
    case 'presentacion':
      icon = <Presentation className="h-6 w-6 text-purple-600" />;
      bgColor = 'bg-purple-50 border-purple-200';
      break;
    case 'ejercicios':
      icon = <FileQuestion className="h-6 w-6 text-amber-600" />;
      bgColor = 'bg-amber-50 border-amber-200';
      break;
  }

  return (
    <div className={`rounded-xl border shadow-sm p-6 h-full flex flex-col ${bgColor}`}>
      <div className="flex items-center justify-between mb-4">
        <div className="flex items-center">
          <div className="rounded-full bg-white p-2">
            {icon}
          </div>
          <h3 className="ml-3 text-lg font-medium text-gray-800">{title}</h3>
        </div>
        {type === 'presentacion' && content && (
          <button
            onClick={() => setShowPreview(!showPreview)}
            className="text-sm text-purple-600 hover:text-purple-800"
          >
            {showPreview ? 'Ocultar Vista Previa' : 'Ver Vista Previa'}
          </button>
        )}
      </div>
      
      {isLoading ? (
        <div className="flex-1 flex justify-center items-center">
          <div className="animate-pulse flex flex-col items-center">
            <div className="h-4 bg-gray-300 rounded w-3/4 mb-2"></div>
            <div className="h-4 bg-gray-300 rounded w-1/2"></div>
          </div>
        </div>
      ) : content ? (
        <>
          <div className="flex-1 overflow-auto bg-white border border-gray-200 rounded-lg p-4 mb-4 text-sm text-gray-700">
            {type === 'presentacion' && showPreview ? (
              renderPresentationPreview()
            ) : (
              <pre className="whitespace-pre-line">{content}</pre>
            )}
          </div>
          
          <div className="flex space-x-2">
            <button 
              onClick={handleCopy} 
              className="flex items-center px-3 py-2 border border-gray-300 rounded-lg text-sm text-gray-700 hover:bg-gray-100 transition-colors duration-200"
            >
              {copied ? <CheckCircle2 className="mr-1 h-4 w-4 text-green-500" /> : <Copy className="mr-1 h-4 w-4" />}
              {copied ? '¡Copiado!' : 'Copiar'}
            </button>
            
            <button 
              onClick={handleDownload} 
              className="flex items-center px-3 py-2 border border-gray-300 rounded-lg text-sm text-gray-700 hover:bg-gray-100 transition-colors duration-200"
            >
              <Download className="mr-1 h-4 w-4" />
              Descargar
            </button>
          </div>
        </>
      ) : (
        <div className="flex-1 flex justify-center items-center text-gray-500 text-sm">
          El contenido generado aparecerá aquí...
        </div>
      )}
    </div>
  );
};

export default ContentDisplay;