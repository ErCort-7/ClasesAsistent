import React, { useState } from 'react';
import { Wand2 } from 'lucide-react';

interface ContenidoClase {
  tema: string;
  materia: string;
  gradoAcademico: string;
  duracion: string;
  tipoClase: 'teorica' | 'practica';
}

interface TopicFormProps {
  onGenerate: (
    tema: string,
    materia: string,
    gradoAcademico: string,
    duracion: string,
    tipoClase: 'teorica' | 'practica'
  ) => void;
  isLoading: boolean;
}

const TopicForm: React.FC<TopicFormProps> = ({ onGenerate, isLoading }) => {
  const [contenido, setContenido] = useState<ContenidoClase>({
    tema: 'Ecuaciones de 1er grado',
    materia: 'matematicas',
    gradoAcademico: '4-primaria',
    duracion: '60',
    tipoClase: 'teorica'
  });

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (contenido.tema.trim()) {
      onGenerate(
        contenido.tema,
        contenido.materia,
        contenido.gradoAcademico,
        contenido.duracion,
        contenido.tipoClase
      );
    }
  };

  const handleInputChange = (
    field: keyof ContenidoClase,
    value: string
  ) => {
    setContenido(prev => ({
      ...prev,
      [field]: value
    }));
  };

  return (
    <div className="bg-white rounded-xl shadow-md p-6 transition-all duration-300 hover:shadow-lg">
      <h2 className="text-xl font-semibold text-gray-800 mb-4">Generar Contenido de Clase</h2>
      
      <form onSubmit={handleSubmit}>
        <div className="mb-6">
          <label htmlFor="tema" className="block text-sm font-medium text-gray-700 mb-1">
            Tema o Concepto
          </label>
          <input
            type="text"
            id="tema"
            value={contenido.tema}
            onChange={(e) => handleInputChange('tema', e.target.value)}
            placeholder="ej: Fracciones, Fotosíntesis, Guerra Mundial..."
            className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-colors duration-200"
            required
          />
        </div>
        
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 mb-6">
          <div>
            <label htmlFor="materia" className="block text-sm font-medium text-gray-700 mb-1">
              Materia
            </label>
            <select
              id="materia"
              value={contenido.materia}
              onChange={(e) => handleInputChange('materia', e.target.value)}
              className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-colors duration-200"
            >
              <option value="matematicas">Matemáticas</option>
              <option value="ciencias">Ciencias Naturales</option>
              <option value="lenguaje">Lengua y Literatura</option>
              <option value="historia">Historia</option>
              <option value="arte">Arte</option>
              <option value="musica">Música</option>
              <option value="educacion-fisica">Educación Física</option>
            </select>
          </div>
          
          <div>
            <label htmlFor="gradoAcademico" className="block text-sm font-medium text-gray-700 mb-1">
              Grado Académico
            </label>
            <select
              id="gradoAcademico"
              value={contenido.gradoAcademico}
              onChange={(e) => handleInputChange('gradoAcademico', e.target.value)}
              className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-colors duration-200"
            >
              <optgroup label="Primaria">
                <option value="1-primaria">1º de Primaria</option>
                <option value="2-primaria">2º de Primaria</option>
                <option value="3-primaria">3º de Primaria</option>
                <option value="4-primaria">4º de Primaria</option>
                <option value="5-primaria">5º de Primaria</option>
                <option value="6-primaria">6º de Primaria</option>
              </optgroup>
              <optgroup label="Secundaria">
                <option value="1-secundaria">1º de Secundaria</option>
                <option value="2-secundaria">2º de Secundaria</option>
                <option value="3-secundaria">3º de Secundaria</option>
              </optgroup>
              <optgroup label="Bachillerato">
                <option value="1-bachillerato">1º de Bachillerato</option>
                <option value="2-bachillerato">2º de Bachillerato</option>
                <option value="3-bachillerato">3º de Bachillerato</option>
              </optgroup>
            </select>
          </div>
          
          <div>
            <label htmlFor="duracion" className="block text-sm font-medium text-gray-700 mb-1">
              Duración de Clase
            </label>
            <select
              id="duracion"
              value={contenido.duracion}
              onChange={(e) => handleInputChange('duracion', e.target.value)}
              className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 transition-colors duration-200"
            >
              <option value="45">45 minutos</option>
              <option value="60">1 hora</option>
              <option value="90">1 hora y media</option>
              <option value="120">2 horas</option>
              <option value="180">3 horas</option>
            </select>
          </div>
        </div>

        <div className="mb-6">
          <label className="block text-sm font-medium text-gray-700 mb-2">
            Tipo de Clase
          </label>
          <div className="flex space-x-4">
            <label className="inline-flex items-center">
              <input
                type="radio"
                className="form-radio text-indigo-600"
                name="tipoClase"
                value="teorica"
                checked={contenido.tipoClase === 'teorica'}
                onChange={(e) => handleInputChange('tipoClase', e.target.value as 'teorica' | 'practica')}
              />
              <span className="ml-2">Teórica</span>
            </label>
            <label className="inline-flex items-center">
              <input
                type="radio"
                className="form-radio text-indigo-600"
                name="tipoClase"
                value="practica"
                checked={contenido.tipoClase === 'practica'}
                onChange={(e) => handleInputChange('tipoClase', e.target.value as 'teorica' | 'practica')}
              />
              <span className="ml-2">Práctica</span>
            </label>
          </div>
        </div>
        
        <button
          type="submit"
          disabled={isLoading || !contenido.tema.trim()}
          className={`w-full flex justify-center items-center px-6 py-3 rounded-lg text-white font-medium ${
            isLoading || !contenido.tema.trim() 
              ? 'bg-indigo-300 cursor-not-allowed' 
              : 'bg-indigo-600 hover:bg-indigo-700 transition-colors duration-200'
          }`}
        >
          {isLoading ? (
            <>
              <div className="animate-spin -ml-1 mr-3 h-5 w-5 text-white border-2 border-white border-t-transparent rounded-full" />
              Generando...
            </>
          ) : (
            <>
              <Wand2 className="mr-2 h-5 w-5" />
              Generar Contenido
            </>
          )}
        </button>
      </form>
    </div>
  );
};

export default TopicForm;