import { createRoot } from 'react-dom/client';
import { App } from './App';
import './style.css';
import './realm.css';
import './workbench.css';
import './reader-compact.css';

createRoot(document.getElementById('root')!).render(<App />);
