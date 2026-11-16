import { createRoot } from 'react-dom/client';
import { App } from './App';
import './style.css';
import './realm.css';
import './workbench.css';
import './reading-layout.css';
import './community.css';

createRoot(document.getElementById('root')!).render(<App />);
