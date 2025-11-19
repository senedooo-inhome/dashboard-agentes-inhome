import React, { createContext, useContext, useState, useEffect } from 'react';

interface AuthContextType {
  token: string | null;
  agentId: string | null;
  login: (token: string, agentId: string) => void;
  logout: () => void;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

export const AuthProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [token, setToken] = useState<string | null>(null);
  const [agentId, setAgentId] = useState<string | null>(null);

  useEffect(() => {
    const storedToken = localStorage.getItem('token');
    const storedAgentId = localStorage.getItem('agentId');
    if (storedToken && storedAgentId) {
      setToken(storedToken);
      setAgentId(storedAgentId);
    }
  }, []);

  const login = (newToken: string, newAgentId: string) => {
    localStorage.setItem('token', newToken);
    localStorage.setItem('agentId', newAgentId);
    setToken(newToken);
    setAgentId(newAgentId);
  };

  const logout = () => {
    localStorage.removeItem('token');
    localStorage.removeItem('agentId');
    setToken(null);
    setAgentId(null);
  };

  return (
    <AuthContext.Provider value={{ token, agentId, login, logout }}>
      {children}
    </AuthContext.Provider>
  );
};

export const useAuth = (): AuthContextType => {
  const context = useContext(AuthContext);
  if (!context) throw new Error('useAuth must be used within AuthProvider');
  return context;
};