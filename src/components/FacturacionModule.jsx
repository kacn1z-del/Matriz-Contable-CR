import React, { useState, useEffect } from 'react';

/**
 * MÓDULO DE FACTURACIÓN - Matriz Contable CR
 * Integra CABYS para facturación electrónica
 * Datos: public/data/cabys-datos.json (20,505 artículos)
 */

const FacturacionModule = () => {
  const [cabysData, setCabysData] = useState([]);
  const [filtrado, setFiltrado] = useState([]);
  const [busqueda, setBusqueda] = useState('');
  const [categoriaFiltro, setCategoriaFiltro] = useState('');
  const [ivaFiltro, setIvaFiltro] = useState('');
  const [articulosSeleccionados, setArticulosSeleccionados] = useState([]);
  const [factura, setFactura] = useState({
    numero: '001',
    fecha: new Date().toISOString().split('T')[0],
    cliente: '',
    cedula_cliente: ''
  });

  // Cargar datos CABYS desde public/data/cabys-datos.json
  useEffect(() => {
    const cargarCabys = async () => {
      try {
        const response = await fetch('/data/cabys-datos.json');
        const datos = await response.json();
        setCabysData(datos);
        setFiltrado(datos);
      } catch (error) {
        console.error('Error cargando CABYS:', error);
      }
    };
    cargarCabys();
  }, []);

  // Filtrar artículos
  useEffect(() => {
    let resultado = cabysData;

    if (busqueda) {
      const termino = busqueda.toLowerCase();
      resultado = resultado.filter(art =>
        art.desc.toLowerCase().includes(termino) ||
        art.codigo.includes(termino)
      );
    }

    if (categoriaFiltro) {
      resultado = resultado.filter(art => art.cat === categoriaFiltro);
    }

    if (ivaFiltro) {
      resultado = resultado.filter(art => art.iva === ivaFiltro);
    }

    setFiltrado(resultado);
  }, [busqueda, categoriaFiltro, ivaFiltro, cabysData]);

  // Agregar artículo a la factura
  const agregarArticulo = (articulo) => {
    const nuevoArticulo = {
      codigo_cabys: articulo.codigo,
      descripcion: articulo.desc,
      cantidad: 1,
      precio_unitario: 0,
      iva: parseInt(articulo.iva, 10) || 0,
      subtotal: 0,
      total_iva: 0,
      total: 0
    };
    setArticulosSeleccionados([...articulosSeleccionados, nuevoArticulo]);
  };

  // Actualizar cantidad o precio y recalcular la línea
  const actualizarArticulo = (index, campo, valor) => {
    const actualizados = [...articulosSeleccionados];
    actualizados[index][campo] = valor;

    if (campo === 'cantidad' || campo === 'precio_unitario') {
      const cantidad = parseFloat(actualizados[index].cantidad) || 0;
      const precio = parseFloat(actualizados[index].precio_unitario) || 0;
      actualizados[index].subtotal = cantidad * precio;
      actualizados[index].total_iva = (actualizados[index].subtotal * actualizados[index].iva) / 100;
      actualizados[index].total = actualizados[index].subtotal + actualizados[index].total_iva;
    }

    setArticulosSeleccionados(actualizados);
  };

  // Eliminar artículo
  const eliminarArticulo = (index) => {
    setArticulosSeleccionados(articulosSeleccionados.filter((_, i) => i !== index));
  };

  // Totales de la factura
  const subtotal = articulosSeleccionados.reduce((sum, art) => sum + art.subtotal, 0);
  const iva_total = articulosSeleccionados.reduce((sum, art) => sum + art.total_iva, 0);
  const total = subtotal + iva_total;

  // Generar factura (aún no timbra en Hacienda)
  const generarFactura = () => {
    if (!factura.cliente || articulosSeleccionados.length === 0) {
      alert('Completa el cliente y agrega al menos un artículo');
      return;
    }

    const facturaCompleta = {
      ...factura,
      articulos: articulosSeleccionados,
      subtotal,
      iva_total,
      total
    };

    console.log('Factura generada:', facturaCompleta);
    alert('Factura generada. Pendiente de timbrar en Hacienda.');
  };

  // Opciones de filtros
  const categorias = [...new Set(cabysData.map(art => art.cat))].sort();
  const ivas_unicos = [...new Set(cabysData.map(art => art.iva))].sort((a, b) => parseInt(a, 10) - parseInt(b, 10));

  return (
    <div style={{ padding: '20px', fontFamily: 'Arial' }}>
      <h1>📋 Facturación Electrónica - Matriz Contable CR</h1>

      {/* DATOS DE LA FACTURA */}
      <div style={{ border: '1px solid #ddd', padding: '15px', marginBottom: '20px', borderRadius: '5px' }}>
        <h3>Datos de la factura</h3>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '10px' }}>
          <input
            type="text"
            placeholder="Número de factura"
            value={factura.numero}
            onChange={(e) => setFactura({ ...factura, numero: e.target.value })}
          />
          <input
            type="date"
            value={factura.fecha}
            onChange={(e) => setFactura({ ...factura, fecha: e.target.value })}
          />
          <input
            type="text"
            placeholder="Nombre del cliente"
            value={factura.cliente}
            onChange={(e) => setFactura({ ...factura, cliente: e.target.value })}
          />
          <input
            type="text"
            placeholder="Cédula/RUC del cliente"
            value={factura.cedula_cliente}
            onChange={(e) => setFactura({ ...factura, cedula_cliente: e.target.value })}
          />
        </div>
      </div>

      {/* BÚSQUEDA DE ARTÍCULOS */}
      <div style={{ border: '1px solid #ddd', padding: '15px', marginBottom: '20px', borderRadius: '5px' }}>
        <h3>🔍 Buscar artículos CABYS</h3>
        <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr 1fr', gap: '10px', marginBottom: '15px' }}>
          <input
            type="text"
            placeholder="Buscar por código o descripción..."
            value={busqueda}
            onChange={(e) => setBusqueda(e.target.value)}
            style={{ padding: '8px' }}
          />
          <select
            value={categoriaFiltro}
            onChange={(e) => setCategoriaFiltro(e.target.value)}
            style={{ padding: '8px' }}
          >
            <option value="">Todas las categorías</option>
            {categorias.map(cat => (
              <option key={cat} value={cat}>{cat}</option>
            ))}
          </select>
          <select
            value={ivaFiltro}
            onChange={(e) => setIvaFiltro(e.target.value)}
            style={{ padding: '8px' }}
          >
            <option value="">Todos los IVA</option>
            {ivas_unicos.map(iva => (
              <option key={iva} value={iva}>{iva}%</option>
            ))}
          </select>
        </div>

        <div style={{ maxHeight: '300px', overflowY: 'auto', border: '1px solid #eee', borderRadius: '3px' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead style={{ backgroundColor: '#f0f0f0', position: 'sticky', top: 0 }}>
              <tr>
                <th style={{ padding: '10px', textAlign: 'left', borderBottom: '2px solid #ddd' }}>Código CABYS</th>
                <th style={{ padding: '10px', textAlign: 'left', borderBottom: '2px solid #ddd' }}>Descripción</th>
                <th style={{ padding: '10px', textAlign: 'center', borderBottom: '2px solid #ddd' }}>IVA</th>
                <th style={{ padding: '10px', textAlign: 'center', borderBottom: '2px solid #ddd' }}>Acción</th>
              </tr>
            </thead>
            <tbody>
              {filtrado.slice(0, 50).map((art) => (
                <tr key={art.codigo} style={{ borderBottom: '1px solid #eee' }}>
                  <td style={{ padding: '10px', fontSize: '12px' }}>{art.codigo}</td>
                  <td style={{ padding: '10px', fontSize: '12px' }}>{art.desc}</td>
                  <td style={{ padding: '10px', textAlign: 'center' }}>{art.iva}%</td>
                  <td style={{ padding: '10px', textAlign: 'center' }}>
                    <button
                      onClick={() => agregarArticulo(art)}
                      style={{
                        padding: '5px 10px',
                        backgroundColor: '#4CAF50',
                        color: 'white',
                        border: 'none',
                        borderRadius: '3px',
                        cursor: 'pointer'
                      }}
                    >
                      Agregar
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <p style={{ padding: '10px', fontSize: '12px', color: '#666' }}>
            Mostrando {Math.min(50, filtrado.length)} de {filtrado.length} artículos
          </p>
        </div>
      </div>

      {/* ARTÍCULOS EN LA FACTURA */}
      <div style={{ border: '1px solid #ddd', padding: '15px', marginBottom: '20px', borderRadius: '5px' }}>
        <h3>📦 Artículos en la factura ({articulosSeleccionados.length})</h3>
        {articulosSeleccionados.length === 0 ? (
          <p style={{ color: '#999' }}>No hay artículos agregados</p>
        ) : (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
              <thead style={{ backgroundColor: '#f0f0f0' }}>
                <tr>
                  <th style={{ padding: '8px', textAlign: 'left', borderBottom: '2px solid #ddd' }}>Descripción</th>
                  <th style={{ padding: '8px', textAlign: 'center', borderBottom: '2px solid #ddd' }}>Cantidad</th>
                  <th style={{ padding: '8px', textAlign: 'right', borderBottom: '2px solid #ddd' }}>Precio unit.</th>
                  <th style={{ padding: '8px', textAlign: 'right', borderBottom: '2px solid #ddd' }}>Subtotal</th>
                  <th style={{ padding: '8px', textAlign: 'center', borderBottom: '2px solid #ddd' }}>IVA %</th>
                  <th style={{ padding: '8px', textAlign: 'right', borderBottom: '2px solid #ddd' }}>IVA ₡</th>
                  <th style={{ padding: '8px', textAlign: 'right', borderBottom: '2px solid #ddd' }}>Total</th>
                  <th style={{ padding: '8px', textAlign: 'center', borderBottom: '2px solid #ddd' }}>Acción</th>
                </tr>
              </thead>
              <tbody>
                {articulosSeleccionados.map((art, idx) => (
                  <tr key={idx} style={{ borderBottom: '1px solid #eee' }}>
                    <td style={{ padding: '8px' }}>{art.descripcion}</td>
                    <td style={{ padding: '8px', textAlign: 'center' }}>
                      <input
                        type="number"
                        min="1"
                        value={art.cantidad}
                        onChange={(e) => actualizarArticulo(idx, 'cantidad', e.target.value)}
                        style={{ width: '60px', padding: '4px' }}
                      />
                    </td>
                    <td style={{ padding: '8px', textAlign: 'right' }}>
                      <input
                        type="number"
                        step="0.01"
                        value={art.precio_unitario}
                        onChange={(e) => actualizarArticulo(idx, 'precio_unitario', e.target.value)}
                        style={{ width: '90px', padding: '4px' }}
                      />
                    </td>
                    <td style={{ padding: '8px', textAlign: 'right' }}>₡{art.subtotal.toFixed(2)}</td>
                    <td style={{ padding: '8px', textAlign: 'center' }}>{art.iva}%</td>
                    <td style={{ padding: '8px', textAlign: 'right' }}>₡{art.total_iva.toFixed(2)}</td>
                    <td style={{ padding: '8px', textAlign: 'right', fontWeight: 'bold' }}>₡{art.total.toFixed(2)}</td>
                    <td style={{ padding: '8px', textAlign: 'center' }}>
                      <button
                        onClick={() => eliminarArticulo(idx)}
                        style={{
                          padding: '4px 8px',
                          backgroundColor: '#f44336',
                          color: 'white',
                          border: 'none',
                          borderRadius: '3px',
                          cursor: 'pointer',
                          fontSize: '11px'
                        }}
                      >
                        Eliminar
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* TOTALES */}
      <div style={{ border: '2px solid #4CAF50', padding: '20px', marginBottom: '20px', borderRadius: '5px', backgroundColor: '#f9f9f9' }}>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '20px', marginBottom: '20px' }}>
          <div>
            <p style={{ margin: '0', color: '#666', fontSize: '12px' }}>Subtotal</p>
            <h3 style={{ margin: '5px 0 0 0', color: '#333' }}>₡{subtotal.toFixed(2)}</h3>
          </div>
          <div>
            <p style={{ margin: '0', color: '#666', fontSize: '12px' }}>Total IVA</p>
            <h3 style={{ margin: '5px 0 0 0', color: '#FF9800' }}>₡{iva_total.toFixed(2)}</h3>
          </div>
          <div>
            <p style={{ margin: '0', color: '#666', fontSize: '12px' }}>Total a pagar</p>
            <h3 style={{ margin: '5px 0 0 0', color: '#4CAF50' }}>₡{total.toFixed(2)}</h3>
          </div>
        </div>

        <button
          onClick={generarFactura}
          style={{
            width: '100%',
            padding: '12px',
            backgroundColor: '#4CAF50',
            color: 'white',
            border: 'none',
            borderRadius: '5px',
            fontSize: '16px',
            fontWeight: 'bold',
            cursor: 'pointer'
          }}
        >
          ✅ Generar factura
        </button>
      </div>
    </div>
  );
};

export default FacturacionModule;
