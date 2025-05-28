import { useCallback } from 'react';

export const useNavigate = () => {
  const navigateTo = useCallback((path: string) => {
    console.log(`Navigation to ${path}`);
  }, []);

  return { navigateTo };
};