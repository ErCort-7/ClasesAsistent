import { useCallback } from 'react';

export const useNavigate = () => {
  const navigateTo = useCallback((path: string) => {
    console.log(`Navigation to ${path}`);
    // In a real app with routing, this would use a router's navigate function
    // For this demo, we're just logging the navigation
  }, []);

  return {
    navigateTo,
  };
};