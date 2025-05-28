import React from 'react';
import HistoryCard from '../components/Dashboard/HistoryCard';
import StatisticsCard from '../components/Dashboard/StatisticsCard';
import { useNavigate } from '../hooks/useNavigate';

const MOCK_HISTORY = [
  {
    id: '1',
    topic: 'Proceso de Fotosíntesis',
    subject: 'Ciencias Naturales',
    gradeLevel: '2º de Secundaria',
    date: '2025-04-01T10:30:00Z',
  },
  {
    id: '2',
    topic: 'Segunda Guerra Mundial: Causas y Efectos',
    subject: 'Historia',
    gradeLevel: '3º de Bachillerato',
    date: '2025-03-28T14:15:00Z',
  },
  {
    id: '3',
    topic: 'Introducción a las Fracciones',
    subject: 'Matemáticas',
    gradeLevel: '4º de Primaria',
    date: '2025-03-25T09:00:00Z',
  },
];

const Dashboard: React.FC = () => {
  const { navigateTo } = useNavigate();
  
  const handleSelectItem = (id: string) => {
    console.log(`Elemento seleccionado: ${id}`);
  };
  
  return (
    <div className="py-8">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="flex items-center justify-between mb-8">
          <h1 className="text-2xl font-bold text-gray-900">Panel Principal</h1>
          <button 
            onClick={() => navigateTo('/generate')}
            className="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg transition-colors duration-200 flex items-center"
          >
            Crear Nuevo Contenido
          </button>
        </div>
        
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          <div className="lg:col-span-2">
            <HistoryCard items={MOCK_HISTORY} onSelectItem={handleSelectItem} />
          </div>
          
          <div>
            <StatisticsCard totalClasses={12} totalPresentations={8} />
          </div>
        </div>
      </div>
    </div>
  );
};

export default Dashboard;