import React from 'react';
import { createRoot } from 'react-dom/client';
import { SpeedInsights } from "@vercel/speed-insights/react";
import App from './App.jsx';
import './styles.css';

createRoot(document.getElementById('root')).render(
  <>
    <App />
    <SpeedInsights />
  </>
);
