import React from 'react';
import { BarChart3, BookOpen, Presentation } from 'lucide-react';

interface StatisticsCardProps {
  totalClasses: number;
  totalPresentations: number;
}

const StatisticsCard: React.FC<StatisticsCardProps> = ({ totalClasses, totalPresentations }) => {
  return (
    <div className="bg-white rounded-xl shadow-md p-6">
      <div className="flex items-center justify-between mb-4">
        <h2 className="text-xl font-semibold text-gray-800">Estadísticas</h2>
        <BarChart3 className="h-5 w-5 text-gray-400" />
      </div>
      
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div className="border border-gray-200 rounded-lg p-4 flex items-center">
          <div className="rounded-full bg-blue-100 p-2">
            <BookOpen className="h-5 w-5 text-blue-600" />
          </div>
          <div className="ml-3">
            <p className="text-sm text-gray-600">Total de Clases</p>
            <p className="text-xl font-semibold text-gray-900">{totalClasses}</p>
          </div>
        </div>
        
        <div className="border border-gray-200 rounded-lg p-4 flex items-center">
          <div className="rounded-full bg-amber-100 p-2">
            <Presentation className="h-5 w-5 text-amber-600" />
          </div>
          <div className="ml-3">
            <p className="text-sm text-gray-600">Presentaciones</p>
            <p className="text-xl font-semibold text-gray-900">{totalPresentations}</p>
          </div>
        </div>
      </div>
    </div>
  );
};

export default StatisticsCard;