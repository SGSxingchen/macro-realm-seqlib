import { createRoot } from 'react-dom/client';
import { App } from './App';
import './style.css';
import './realm.css';
import './workbench.css';
import './reading-layout.css';

createRoot(document.getElementById('root')!).render(<App />);
