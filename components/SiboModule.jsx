import React, { useState } from 'react'
import './SiboModule.css'

export default function SiboModule({ i18n }) {
  const [bienes, setBienes] = useState([])
  const [formData, setFormData] = useState({
    descripcion: '',
    valor: '',
    fecha: new Date().toISOString().split('T')[0]
  })
  const [mensaje, setMensaje] = useState('')

  const calcularPenalizacion = (valor) => {
    return parseFloat(valor) * 0.30
  }

  const calcularIVA = (valor) => {
    return parseFloat(valor) * 0.13
  }

  const agregarBien = () => {
    if (!formData.descripcion || !formData.valor) {
      setMensaje('⚠️ Completa todos los campos')
      return
    }

    const penalizacion = calcularPenalizacion(formData.valor)
    const iva = calcularIVA(formData.valor)

    const nuevoBien = {
      id: Date.now(),
      descripcion: formData.descripcion,
      valor: parseFloat(formData.valor),
      penalizacion: penalizacion,
      iva: iva,
      total: parseFloat(formData.valor) + penalizacion + iva,
      fecha: formData.fecha,
      estado: 'Registrado'
    }

    setBienes([...bienes, nuevoBien])
    setFormData({ descripcion: '', valor: '', fecha: new Date().toISOString().split('T')[0] })
    setMensaje('✅ Bien omitido registrado')
    setTimeout(() => setMensaje(''), 3000)
  }

  const totalValor = bienes.reduce((sum, b) => sum + b.valor, 0)
  const totalPenalizacion = bienes.reduce((sum, b) => sum + b.penalizacion, 0)
  const totalIVA = bienes.reduce((sum, b) => sum + b.iva, 0)
  const totalAdeudo = totalValor + totalPenalizacion + totalIVA

  const eliminarBien = (id) => {
    setBienes(bienes.filter(b => b.id !== id))
  }

  const formatCurrency = (value) => {
    return parseFloat(value).toLocaleString('es-CR', {
      style: 'currency',
      currency: 'CRC',
      minimumFractionDigits: 2
    })
  }

  return (
    <div className="sibo-module">
      <div className="sibo-header">
        <h1>📊 SIBO - Bienes Omitidos</h1>
        <p>Sistema de Información de Bienes Omitidos para Costa Rica</p>
      </div>

      {mensaje && <div className="mensaje">{mensaje}</div>}

      <div className="stats-grid">
        <div className="stat-card">
          <div className="stat-icon">💰</div>
          <div className="stat-label">Valor Total</div>
          <div className="stat-valor">{formatCurrency(totalValor)}</div>
        </div>
        <div className="stat-card">
          <div className="stat-icon">⚠️</div>
          <div className="stat-label">Penalización (30%)</div>
          <div className="stat-valor">{formatCurrency(totalPenalizacion)}</div>
        </div>
        <div className="stat-card">
          <div className="stat-icon">📈</div>
          <div className="stat-label">IVA (13%)</div>
          <div className="stat-valor">{formatCurrency(totalIVA)}</div>
        </div>
        <div className="stat-card">
          <div className="stat-icon">💳</div>
          <div className="stat-label">Total Adeudo</div>
          <div className="stat-valor">{formatCurrency(totalAdeudo)}</div>
        </div>
      </div>

      <div className="sibo-form">
        <h2>Registrar Bien Omitido</h2>
        <div className="form-group">
          <label>Descripción del Bien</label>
          <input
            type="text"
            placeholder="Ej: Maquinaria, Vehículo, etc."
            value={formData.descripcion}
            onChange={(e) => setFormData({...formData, descripcion: e.target.value})}
          />
        </div>
        <div className="form-row">
          <div className="form-group">
            <label>Valor (₡)</label>
            <input
              type="number"
              placeholder="0.00"
              value={formData.valor}
              onChange={(e) => setFormData({...formData, valor: e.target.value})}
            />
          </div>
          <div className="form-group">
            <label>Fecha de Omisión</label>
            <input
              type="date"
              value={formData.fecha}
              onChange={(e) => setFormData({...formData, fecha: e.target.value})}
            />
          </div>
        </div>
        <button className="btn-agregar" onClick={agregarBien}>Registrar Bien ✓</button>
      </div>

      <div className="sibo-listado">
        <h2>Bienes Registrados ({bienes.length})</h2>
        {bienes.length === 0 ? (
          <p className="sin-datos">No hay bienes registrados</p>
        ) : (
          <table>
            <thead>
              <tr>
                <th>Descripción</th>
                <th>Valor</th>
                <th>Penalización (30%)</th>
                <th>IVA (13%)</th>
                <th>Total</th>
                <th>Fecha</th>
                <th>Acción</th>
              </tr>
            </thead>
            <tbody>
              {bienes.map(b => (
                <tr key={b.id}>
                  <td>{b.descripcion}</td>
                  <td>{formatCurrency(b.valor)}</td>
                  <td className="penalizacion">{formatCurrency(b.penalizacion)}</td>
                  <td>{formatCurrency(b.iva)}</td>
                  <td className="total">{formatCurrency(b.total)}</td>
                  <td>{b.fecha}</td>
                  <td>
                    <button className="btn-delete" onClick={() => eliminarBien(b.id)}>✕</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      <div className="sibo-resumen">
        <h3>Resumen de Cálculos</h3>
        <div className="resumen-grid">
          <div className="resumen-item">
            <div className="resumen-label">Valor de Bienes</div>
            <div className="resumen-valor">{formatCurrency(totalValor)}</div>
          </div>
          <div className="resumen-item">
            <div className="resumen-label">Penalización</div>
            <div className="resumen-valor">{formatCurrency(totalPenalizacion)}</div>
          </div>
          <div className="resumen-item">
            <div className="resumen-label">IVA Adeudado</div>
            <div className="resumen-valor">{formatCurrency(totalIVA)}</div>
          </div>
          <div className="resumen-item total">
            <div className="resumen-label">Total Adeudo</div>
            <div className="resumen-valor">{formatCurrency(totalAdeudo)}</div>
          </div>
        </div>
      </div>
    </div>
  )
}
