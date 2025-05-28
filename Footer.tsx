import React from 'react';
import { Heart } from 'lucide-react';

const Footer: React.FC = () => {
  return (
    <footer className="bg-gray-50 border-t border-gray-200">
      <div className="max-w-7xl mx-auto py-8 px-4 sm:px-6 lg:px-8">
        <div className="md:flex md:items-center md:justify-between">
          <div className="flex justify-center md:justify-start">
            <p className="text-center text-sm text-gray-500">
              &copy; {new Date().getFullYear()} EduGenius. Todos los derechos reservados.
            </p>
          </div>
          <div className="flex items-center justify-center mt-4 md:mt-0">
            <span className="text-sm text-gray-500 flex items-center">
              Hecho con <Heart className="h-4 w-4 text-red-500 mx-1" /> para docentes
            </span>
          </div>
        </div>
        <div className="mt-4 flex justify-center md:justify-start space-x-6">
          <a href="#" className="text-gray-500 hover:text-gray-700 transition-colors duration-200">
            Política de Privacidad
          </a>
          <a href="#" className="text-gray-500 hover:text-gray-700 transition-colors duration-200">
            Términos de Servicio
          </a>
          <a href="#" className="text-gray-500 hover:text-gray-700 transition-colors duration-200">
            Contacto
          </a>
        </div>
      </div>
    </footer>
  );
};

export default Footer;