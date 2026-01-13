import React from 'react';
import { BrowserRouter as Router, Routes, Route } from 'react-router-dom';
import Login from './components/Login';
import Dashboard from './components/Dashboard';
import Sheet from './components/Sheet';
import Settings from './components/Settings';

function App() {
  return (
    <Router>
      <Routes>
        <Route path="/" element={<Login />} />
        <Route path="/dashboard" element={<Dashboard />} />
        <Route path="/sheet/:id" element={<Sheet />} />
        <Route path="/settings/:id" element={<Settings />} />
      </Routes>
    </Router>
  );
}

export default App;
