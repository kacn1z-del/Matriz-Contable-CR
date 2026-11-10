/**
 * SERVICIO CABYS
 * Gestión de búsqueda y filtrado del Catálogo de Bienes y Servicios
 * Matriz Contable CR v37.5
 */

class CabyService {
  constructor() {
    this.cabysData = [];
    this.categorias = [];
    this.ivas = [];
  }

  /**
   * Cargar datos CABYS desde API o archivo
   */
  async cargarDatos(origen = 'local') {
    try {
      let datos;
      
      if (origen === 'api') {
        // Desde API backend
        const response = await fetch('/api/cabys');
        datos = await response.json();
      } else if (origen === 'comprimido') {
        // Desde archivo GZ (necesita decompresión)
        const response = await fetch('/data/cabys-datos.json.gz');
        const buffer = await response.arrayBuffer();
        // Usar pako para descomprimir
        const decompressed = pako.ungzip(new Uint8Array(buffer), { to: 'string' });
        datos = JSON.parse(decompressed);
      } else {
        // Desde archivo JSON local
        const response = await fetch('/data/cabys-datos.json');
        datos = await response.json();
      }
      
      this.cabysData = datos;
      this.extraerCategorias();
      this.extraerIvas();
      
      console.log(`✅ CABYS cargado: ${this.cabysData.length} artículos`);
      return true;
    } catch (error) {
      console.error('❌ Error cargando CABYS:', error);
      return false;
    }
  }

  /**
   * Extraer categorías únicas
   */
  extraerCategorias() {
    this.categorias = [...new Set(this.cabysData.map(item => item.cat))].sort();
  }

  /**
   * Extraer tasas de IVA únicas
   */
  extraerIvas() {
    this.ivas = [...new Set(this.cabysData.map(item => item.iva))].sort((a, b) => {
      const aNum = parseInt(a);
      const bNum = parseInt(b);
      return aNum - bNum;
    });
  }

  /**
   * Buscar artículos por múltiples criterios
   * @param {string} busqueda - Término de búsqueda (código o descripción)
   * @param {string} categoria - Categoría para filtrar
   * @param {string} iva - Tasa de IVA para filtrar
   * @param {number} limite - Máximo de resultados
   */
  buscar(busqueda = '', categoria = '', iva = '', limite = 100) {
    let resultado = this.cabysData;

    // Filtro por búsqueda (código o descripción)
    if (busqueda.trim()) {
      const termino = busqueda.toLowerCase();
      resultado = resultado.filter(item =>
        item.codigo.includes(termino) ||
        item.desc.toLowerCase().includes(termino)
      );
    }

    // Filtro por categoría
    if (categoria) {
      resultado = resultado.filter(item => item.cat === categoria);
    }

    // Filtro por IVA
    if (iva) {
      resultado = resultado.filter(item => item.iva === iva);
    }

    // Limitar resultados
    return resultado.slice(0, limite);
  }

  /**
   * Obtener artículo por código CABYS
   */
  obtenerPorCodigo(codigo) {
    return this.cabysData.find(item => item.codigo === codigo);
  }

  /**
   * Obtener artículos por categoría
   */
  obtenerPorCategoria(categoria, limite = 100) {
    return this.cabysData
      .filter(item => item.cat === categoria)
      .slice(0, limite);
  }

  /**
   * Obtener artículos por tasa de IVA
   */
  obtenerPorIva(iva, limite = 100) {
    return this.cabysData
      .filter(item => item.iva === iva)
      .slice(0, limite);
  }

  /**
   * Obtener categorías
   */
  getCategorias() {
    return this.categorias;
  }

  /**
   * Obtener tasas de IVA
   */
  getIvas() {
    return this.ivas;
  }

  /**
   * Obtener estadísticas
   */
  getEstadisticas() {
    const stats = {
      totalArticulos: this.cabysData.length,
      totalCategorias: this.categorias.length,
      totalIvas: this.ivas.length,
      articulosPorIva: {}
    };

    // Contar artículos por IVA
    this.ivas.forEach(iva => {
      stats.articulosPorIva[iva] = this.cabysData.filter(item => item.iva === iva).length;
    });

    return stats;
  }

  /**
   * Generar JSON para factura
   */
  generarLineaFactura(codigo, cantidad, precioUnitario) {
    const articulo = this.obtenerPorCodigo(codigo);
    if (!articulo) throw new Error(`Artículo ${codigo} no encontrado`);

    const iva = parseInt(articulo.iva);
    const subtotal = cantidad * precioUnitario;
    const totalIva = subtotal * (iva / 100);
    const total = subtotal + totalIva;

    return {
      codigo_cabys: articulo.codigo,
      descripcion: articulo.desc,
      cantidad,
      precio_unitario: precioUnitario,
      iva,
      subtotal: parseFloat(subtotal.toFixed(2)),
      total_iva: parseFloat(totalIva.toFixed(2)),
      total: parseFloat(total.toFixed(2))
    };
  }

  /**
   * Validar código CABYS
   */
  validarCodigo(codigo) {
    return this.cabysData.some(item => item.codigo === codigo);
  }

  /**
   * Exportar datos a CSV
   */
  exportarCSV() {
    let csv = 'Código CABYS,Descripción,IVA,Categoría,Ruta\n';
    this.cabysData.forEach(item => {
      csv += `"${item.codigo}","${item.desc}","${item.iva}%","${item.cat}","${item.ruta}"\n`;
    });
    return csv;
  }

  /**
   * Descargar archivo CSV
   */
  descargarCSV(nombreArchivo = 'cabys-export.csv') {
    const csv = this.exportarCSV();
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = nombreArchivo;
    link.click();
  }
}

// Crear instancia global
const cabyService = new CabyService();

// Exportar para uso en componentes
export default cabyService;
