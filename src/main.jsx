import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './minimal/App.jsx';
import './index.css';

// Wait for Office.js to be ready before rendering the app
Office.onReady((info) => {
  console.log('Office.js ready, initializing React app');
  console.log('Host:', info.host);
  console.log('Platform:', info.platform);
  
  if (info.host === Office.HostType.Word) {
    console.log('? Running in Word');
  } else {
    console.warn('?? Not running in Word, host is:', info.host);
  }
  
  // Render the React app only after Office.js is ready
  ReactDOM.createRoot(document.getElementById('root')).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
}).catch((error) => {
  console.error('? Office.js failed to initialize:', error);
  
  // Show error message in the UI
  const root = document.getElementById('root');
  root.innerHTML = `
    <div style="padding: 20px; color: red; font-family: sans-serif;">
      <h2>Office.js Initialization Error</h2>
      <p>Failed to initialize Office.js. Please ensure you're running this add-in in Microsoft Word.</p>
      <pre>${error.message || error}</pre>
    </div>
  `;
});

