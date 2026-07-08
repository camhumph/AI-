import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
// @ts-expect-error -- css-only package ships no type declarations
import '@fontsource-variable/inter'
import './index.css'
import App from './App.tsx'

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
)
