import React, { useState } from 'react'
import './App.css'

// Componentes Nuevos
import FacturaElectronica from './components/FacturaElectronica'
import SiboModule from './components/SiboModule'
import FormulasModule from './components/FormulasModule'

// Mock i18n
const i18n = {
  t: (key) => key,
  lang: 'es'
}

function App() {
  const [activePage, setActivePage] = useState('inicio')
  const [sidebarOpen, setSidebarOpen] = useState(true)

  return (
    <div className="app-container">
      {/* HEADER */}
      <header className="app-header">
        <div className="header-left">
          <button
            className="btn-toggle-sidebar"
            onClick={() => setSidebarOpen(!sidebarOpen)}
          >
            ☰
          </button>
          <h1 className="app-title">📊 Matriz Contable CR v39.5+</h1>
        </div>
        <div className="header-right">
          <span className="user-info">👤 Usuario</span>
        </div>
      </header>

      <div className="app-layout">
        {/* SIDEBAR */}
        <nav className={`app-sidebar ${sidebarOpen ? 'open' : 'closed'}`}>
          <div className="nav-section">
            <h3 className="nav-title">💼 Finanzas</h3>
            <button
              className={`nav-btn ${activePage === 'factura' ? 'active' : ''}`}
              onClick={() => setActivePage('factura')}
            >
              📄 Factura Electrónica
            </button>
          </div>

          <div className="nav-section">
            <h3 className="nav-title">🎯 Costa Rica</h3>
            <button
              className={`nav-btn ${activePage === 'sibo' ? 'active' : ''}`}
              onClick={() => setActivePage('sibo')}
            >
              📊 Sibo (Bienes Omitidos)
            </button>
          </div>

          <div className="nav-section">
            <h3 className="nav-title">🛠️ Herramientas</h3>
            <button
              className={`nav-btn ${activePage === 'formulas' ? 'active' : ''}`}
              onClick={() => setActivePage('formulas')}
            >
              📐 510+ Fórmulas
            </button>
          </div>
        </nav>

        {/* MAIN CONTENT */}
        <main className="app-main">
          {/* Factura Electrónica */}
          {activePage === 'factura' && (
            <FacturaElectronica i18n={i18n} />
          )}

          {/* Sibo */}
          {activePage === 'sibo' && (
            <SiboModule i18n={i18n} />
          )}

          {/* 510+ Fórmulas */}
          {activePage === 'formulas' && (
            <FormulasModule i18n={i18n} />
          )}

          {/* Página de Inicio */}
          {activePage === 'inicio' && (
            <div className="inicio-page">
              <div className="welcome-card">
                <h1>📊 Bienvenido a Matriz Contable CR v39.5+</h1>
                <p>Sistema completo de contabilidad para Costa Rica</p>
                <div className="features-grid">
                  <div className="feature-card">
                    <div className="feature-icon">📄</div>
                    <h3>Factura Electrónica</h3>
                    <p>Emisión de facturas electrónicas para Hacienda CR</p>
                  </div>
                  <div className="feature-card">
                    <div className="feature-icon">📊</div>
                    <h3>SIBO</h3>
                    <p>Registro de bienes omitidos con cálculo automático</p>
                  </div>
                  <div className="feature-card">
                    <div className="feature-icon">📐</div>
                    <h3>510+ Fórmulas</h3>
                    <p>Librería completa de funciones financieras</p>
                  </div>
                </div>
              </div>
            </div>
          )}
        </main>
      </div>

      {/* FOOTER */}
      <footer className="app-footer">
        <p>© 2026 Matriz Contable CR | Desarrollado con React v39.5+</p>
        <p>Última actualización: Octubre 7, 2026</p>
      </footer>
    </div>
  )
}

export default App
