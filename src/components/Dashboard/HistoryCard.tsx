import React from 'react';
import { Clock, ArrowRight } from 'lucide-react';

interface HistoryItem {
  id: string;
  topic: string;
  subject: string;
  gradeLevel: string;
  date: string;
}

interface HistoryCardProps {
  items: HistoryItem[];
  onSelectItem: (id: string) => void;
}

const HistoryCard: React.FC<HistoryCardProps> = ({ items, onSelectItem }) => {
  const formatDate = (dateStr: string) => {
    const date = new Date(dateStr);
    return new Intl.DateTimeFormat('es-ES', {
      day: 'numeric',
      month: 'long',
      year: 'numeric',
    }).format(date);
  };

  return (
    <div className="bg-white rounded-xl shadow-md p-6">
      <div className="flex items-center justify-between mb-4">
        <h2 className="text-xl font-semibold text-gray-800">Clases Recientes</h2>
        <Clock className="h-5 w-5 text-gray-400" />
      </div>
      
      {items.length > 0 ? (
        <div className="space-y-4">
          {items.map((item) => (
            <div
              key={item.id}
              className="border border-gray-200 rounded-lg p-4 hover:bg-gray-50 transition-colors duration-200 cursor-pointer"
              onClick={() => onSelectItem(item.id)}
            >
              <div className="flex justify-between items-start">
                <div>
                  <h3 className="font-medium text-indigo-600">{item.topic}</h3>
                  <p className="text-sm text-gray-600">
                    {item.subject} • {item.gradeLevel}
                  </p>
                </div>
                <span className="text-xs text-gray-500">{formatDate(item.date)}</span>
              </div>
              <div className="mt-2 flex justify-end">
                <button className="text-sm text-indigo-600 hover:text-indigo-800 flex items-center transition-colors duration-200">
                  Ver <ArrowRight className="ml-1 h-3 w-3" />
                </button>
              </div>
            </div>
          ))}
        </div>
      ) : (
        <div className="text-center py-6 text-gray-500">
          <p>No hay clases recientes.</p>
          <p className="text-sm mt-1">¡Genera tu primera clase para verla aquí!</p>
        </div>
      )}
    </div>
  );
};

export default HistoryCard;