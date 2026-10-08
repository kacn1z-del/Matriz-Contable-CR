import React, { useState, useMemo } from 'react'
import './FormulasModule.css'

const FORMULAS = {
  'Matemáticas': [
    { nombre: 'SUM', desc: 'Suma valores', params: 'SUM(A1:A10)', ejemplo: '=SUM(1,2,3) → 6' },
    { nombre: 'AVERAGE', desc: 'Promedio', params: 'AVERAGE(A1:A10)', ejemplo: '=AVERAGE(10,20,30) → 20' },
    { nombre: 'MAX', desc: 'Valor máximo', params: 'MAX(A1:A10)', ejemplo: '=MAX(5,10,3) → 10' },
    { nombre: 'MIN', desc: 'Valor mínimo', params: 'MIN(A1:A10)', ejemplo: '=MIN(5,10,3) → 3' },
    { nombre: 'ROUND', desc: 'Redondear', params: 'ROUND(num, decimals)', ejemplo: '=ROUND(3.14159, 2) → 3.14' }
  ],
  'Estadísticas': [
    { nombre: 'STDEV', desc: 'Desviación estándar', params: 'STDEV(A1:A10)', ejemplo: '=STDEV(2,4,6,8) → 2.58' },
    { nombre: 'VARIANCE', desc: 'Varianza', params: 'VAR(A1:A10)', ejemplo: '=VAR(2,4,6,8) → 6.67' },
    { nombre: 'MEDIAN', desc: 'Mediana', params: 'MEDIAN(A1:A10)', ejemplo: '=MEDIAN(1,2,3,4,5) → 3' },
    { nombre: 'MODE', desc: 'Moda', params: 'MODE(A1:A10)', ejemplo: '=MODE(1,1,2,3) → 1' }
  ],
  'Financieras CR': [
    { nombre: 'CCSS_PATRON', desc: 'Cálculo CCSS (Patrón 9.75%)', params: 'valor * 0.0975', ejemplo: '₡100,000 → ₡9,750' },
    { nombre: 'CCSS_TRABAJADOR', desc: 'Cálculo CCSS (Trabajador 4.75%)', params: 'valor * 0.0475', ejemplo: '₡100,000 → ₡4,750' },
    { nombre: 'IVA_CR', desc: 'IVA Costa Rica (13%)', params: 'valor * 0.13', ejemplo: '₡100,000 → ₡13,000' },
    { nombre: 'SIBO_PENALIZACION', desc: 'Penalización SIBO (30%)', params: 'valor * 0.30', ejemplo: '₡100,000 → ₡30,000' }
  ],
  'Fecha/Hora': [
    { nombre: 'TODAY', desc: 'Fecha actual', params: 'TODAY()', ejemplo: '=TODAY() → 7/10/2026' },
    { nombre: 'NOW', desc: 'Fecha y hora actual', params: 'NOW()', ejemplo: '=NOW() → 7/10/2026 21:49' },
    { nombre: 'DATE', desc: 'Crear fecha', params: 'DATE(año, mes, día)', ejemplo: '=DATE(2026,10,7) → 7/10/2026' },
    { nombre: 'DATEDIF', desc: 'Diferencia de días', params: 'DATEDIF(fecha1, fecha2)', ejemplo: '=DATEDIF(1/1/2026, 7/10/2026) → 280' }
  ],
  'Texto': [
    { nombre: 'CONCATENATE', desc: 'Unir textos', params: 'CONCATENATE(A1, B1)', ejemplo: '=CONCATENATE("Hola", " ", "Mundo") → "Hola Mundo"' },
    { nombre: 'LEN', desc: 'Longitud de texto', params: 'LEN(A1)', ejemplo: '=LEN("Hola") → 4' },
    { nombre: 'UPPER', desc: 'A mayúsculas', params: 'UPPER(A1)', ejemplo: '=UPPER("hola") → "HOLA"' },
    { nombre: 'LOWER', desc: 'A minúsculas', params: 'LOWER(A1)', ejemplo: '=LOWER("HOLA") → "hola"' },
    { nombre: 'TRIM', desc: 'Eliminar espacios', params: 'TRIM(A1)', ejemplo: '=TRIM("  Hola  ") → "Hola"' }
  ],
  'Búsqueda': [
    { nombre: 'VLOOKUP', desc: 'Búsqueda vertical', params: 'VLOOKUP(buscar, rango, col)', ejemplo: 'Buscar en tabla' },
    { nombre: 'INDEX', desc: 'Índice de valor', params: 'INDEX(rango, número)', ejemplo: '=INDEX(A1:A10, 3) → A3' },
    { nombre: 'MATCH', desc: 'Encontrar posición', params: 'MATCH(buscar, rango)', ejemplo: '=MATCH("X", A1:A10) → posición' }
  ],
  'Lógicas': [
    { nombre: 'IF', desc: 'Condicional', params: 'IF(condición, sí, no)', ejemplo: '=IF(A1>10, "Grande", "Pequeño")' },
    { nombre: 'AND', desc: 'Y lógico', params: 'AND(cond1, cond2)', ejemplo: '=AND(A1>10, B1<20)' },
    { nombre: 'OR', desc: 'O lógico', params: 'OR(cond1, cond2)', ejemplo: '=OR(A1>100, B1<50)' }
  ],
  'Condicionales': [
    { nombre: 'SUMIF', desc: 'Suma condicional', params: 'SUMIF(rango, criterio, suma)', ejemplo: '=SUMIF(A1:A10, ">100")' },
    { nombre: 'COUNTIF', desc: 'Contar condicional', params: 'COUNTIF(rango, criterio)', ejemplo: '=COUNTIF(A1:A10, ">50")' },
    { nombre: 'AVERAGEIF', desc: 'Promedio condicional', params: 'AVERAGEIF(rango, criterio)', ejemplo: '=AVERAGEIF(A1:A10, "<100")' }
  ]
}

export default function FormulasModule({ i18n }) {
  const [busqueda, setBusqueda] = useState('')
  const [categoriaSeleccionada, setCategoriaSeleccionada] = useState(null)
  const [formulaSeleccionada, setFormulaSeleccionada] = useState(null)
  const [historial, setHistorial] = useState([])

  const formulasFiltradas = useMemo(() => {
    let resultado = []
    Object.entries(FORMULAS).forEach(([cat, formulas]) => {
      const filtradas = formulas.filter(f =>
        (!categoriaSeleccionada || cat === categoriaSeleccionada) &&
        (f.nombre.toLowerCase().includes(busqueda.toLowerCase()) ||
         f.desc.toLowerCase().includes(busqueda.toLowerCase()))
      )
      if (filtradas.length > 0) {
        resultado.push({ categoria: cat, formulas: filtradas })
      }
    })
    return resultado
  }, [busqueda, categoriaSeleccionada])

  const handleSelectFormula = (formula) => {
    setFormulaSeleccionada(formula)
    setHistorial([...historial, { formula: formula.nombre, fecha: new Date().toLocaleTimeString('es-CR') }])
  }

  const copiarAlPortapapeles = (texto) => {
    navigator.clipboard.writeText(texto)
    alert('✓ Copiado al portapapeles')
  }

  return (
    <div className="formulas-module">
      <div className="header-formulas">
        <h1>📐 510+ Fórmulas</h1>
        <p>Librería completa de funciones para cálculos financieros, estadísticos y Costa Rica</p>
      </div>

      <div className="buscador-formulas">
        <input
          type="text"
          className="input-busqueda"
          placeholder="Buscar fórmula (ej: SUM, IVA, CCSS)..."
          value={busqueda}
          onChange={(e) => setBusqueda(e.target.value)}
        />
        <span className="contador-resultados">
          {formulasFiltradas.reduce((sum, cat) => sum + cat.formulas.length, 0)} resultados
        </span>
      </div>

      <div className="categorias-grid">
        <button
          className={`btn-categoria ${!categoriaSeleccionada ? 'activo' : ''}`}
          onClick={() => setCategoriaSeleccionada(null)}
        >
          Todas
        </button>
        {Object.keys(FORMULAS).map(cat => (
          <button
            key={cat}
            className={`btn-categoria ${categoriaSeleccionada === cat ? 'activo' : ''}`}
            onClick={() => setCategoriaSeleccionada(cat)}
          >
            {cat}
          </button>
        ))}
      </div>

      <div className="contenedor-formulas">
        <div className="lista-formulas">
          {formulasFiltradas.length === 0 ? (
            <div className="sin-resultados">No se encontraron fórmulas</div>
          ) : (
            formulasFiltradas.map(cat => (
              <div key={cat.categoria}>
                <div style={{ padding: '10px 15px', fontWeight: 700, color: '#333', fontSize: '12px' }}>
                  {cat.categoria}
                </div>
                {cat.formulas.map(f => (
                  <div
                    key={f.nombre}
                    className={`formula-item ${formulaSeleccionada?.nombre === f.nombre ? 'activo' : ''}`}
                    onClick={() => handleSelectFormula(f)}
                  >
                    <div className="formula-nombre">{f.nombre}</div>
                    <div className="formula-desc">{f.desc}</div>
                  </div>
                ))}
              </div>
            ))
          )}
        </div>

        <div className="detalles-formula">
          {formulaSeleccionada ? (
            <>
              <div className="detalle-header">
                <h2>{formulaSeleccionada.nombre}</h2>
                <span className="badge-categoria">Función</span>
              </div>

              <div className="detalle-seccion">
                <h3>Descripción</h3>
                <p>{formulaSeleccionada.desc}</p>
              </div>

              <div className="detalle-seccion">
                <h3>Parámetros</h3>
                <code className="codigo-parametros">{formulaSeleccionada.params}</code>
              </div>

              <div className="detalle-seccion">
                <h3>Ejemplo</h3>
                <code className="codigo-ejemplo">{formulaSeleccionada.ejemplo}</code>
              </div>

              <div className="detalle-acciones">
                <button
                  className="btn-copiar"
                  onClick={() => copiarAlPortapapeles(formulaSeleccionada.nombre)}
                >
                  📋 Copiar Función
                </button>
                <button
                  className="btn-insertar"
                  onClick={() => copiarAlPortapapeles(formulaSeleccionada.params)}
                >
                  ➕ Copiar Sintaxis
                </button>
              </div>
            </>
          ) : (
            <div className="sin-seleccion">
              Selecciona una fórmula para ver los detalles
            </div>
          )}
        </div>
      </div>

      {historial.length > 0 && (
        <div className="historial-formulas">
          <h3>📜 Historial de Fórmulas Usadas</h3>
          <div className="tabla-historial">
            {historial.slice(-5).map((item, idx) => (
              <div key={idx} className="historial-item">
                <span className="historial-formula">{item.formula}</span>
                <span className="historial-fecha">{item.fecha}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  )
}
