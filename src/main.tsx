import React from 'react'
import ReactDOM from 'react-dom/client'
import SiteReport from './SiteApp' // Import the new component
import './style.css'

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <SiteReport /> {/* Render the Site Report Builder */}
  </React.StrictMode>,
)
