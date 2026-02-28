import React from 'react';
import { BrowserRouter as Router, Routes, Route } from 'react-router-dom';
import Login from './components/Login';
import Dashboard from './components/Dashboard';
import Projects from './components/Projects';
import Sheet from './components/Sheet';
import DataSheet from './components/DataSheet';
import Document from './components/Document';
import Settings from './components/Settings';
import ChangePassword from './components/ChangePassword';
import Admin from './components/Admin';

function App() {
  return (
    <Router>
      <Routes>
        <Route path="/" element={<Login />} />
        <Route path="/projects" element={<Projects />} />
        <Route path="/project/:project" element={<Dashboard />} />
        <Route path="/sheet/:id" element={<DataSheet />} />
        <Route path="/document/:id" element={<Document />} />
        <Route path="/settings/:id" element={<Settings />} />
        <Route path="/change-password" element={<ChangePassword />} />
        <Route path="/admin" element={<Admin />} />
      </Routes>
    </Router>
  );
}

export default App;
