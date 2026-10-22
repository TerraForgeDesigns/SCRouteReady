import React from 'react';
import ReactDOM from 'react-dom/client';
import { HubRouteApp } from './integration/HubRouteApp';
import './styles/app.css';
import './styles/pull-sheet.css';
import './styles/print.css';
ReactDOM.createRoot(document.getElementById('root')!).render(<React.StrictMode><HubRouteApp/></React.StrictMode>);
