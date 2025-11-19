import React from 'react';
import LoginForm from './components/LoginForm';
import Dashboard from './components/Dashboard';
import { AuthProvider, useAuth } from './src/contexts/AuthContext';

const MainApp: React.FC = () => {
  const { token, agentId, login, logout } = useAuth();

  return (
    <>
      {token && agentId ? (
        <Dashboard token={token} agentId={agentId} onLogout={logout} />
      ) : (
        <LoginForm onLoginSuccess={login} />
      )}
    </>
  );
};

const App: React.FC = () => {
  return (
    <AuthProvider>
      <MainApp />
    </AuthProvider>
  );
};

export default App;