# Matriz-Contable-CR
# Matriz Contable CR v12 🇨🇷

**App contable offline-first para Costa Rica** con factura electrónica 4.3, canvas libre, 300+ funciones de hoja de cálculo y asistente Sibö ✦ en 10 lenguas costarricenses.

## Descripción

Aplicación web de una sola página que funciona 100% offline. Pensada para pymes, contadores y emprendedores ticos que necesitan llevar libros contables, emitir facturas para Hacienda y hacer reportes sin depender de internet o suscripciones.

## Características principales

### 1. Factura Electrónica 4.3
- Genera XML + PDF con QR para Ministerio de Hacienda CR
- IVA configurable: 13% general, 4% canasta básica, 2% medicinas, 1% equipo médico, 0% exento
- Emisor/Receptor editables: nombre, cédula, teléfono, email
- Líneas de detalle dinámicas con cálculo automático de subtotal, IVA y total
- Exporta XML 4.3 listo para enviar a Hacienda
- Exporta PDF para impresión/entrega cliente

### 2. Canvas Libre Contable
- Área de trabajo con cuadrícula para crear documentos contables
- Tablas editables directamente en las celdas
- 6 Plantillas incluidas: Libro Diario, Libro Mayor, Balance de Comprobación, Conciliación Bancaria, Factura CR, Reporte Compras/Ventas D-151
- Guardado automático en `localStorage`
- Exportar canvas a XML

### 3. Biblioteca de Funciones - 300+
Modal con 12 categorías de fórmulas tipo hoja de cálculo:

| Categoría | Funciones destacadas |
| --- | --- |
| **Búsqueda y Referencia** | BUSCARV, BUSCARX, INDICE, COINCIDIR, FILTRAR, ORDENAR, UNICOS |
| **Texto** | CONCAT, EXTRAE, SUSTITUIR, DIVIDIRTEXTO, REGEXEXTRACCION, TRADUCIR |
| **Lógicas** | SI, SI.CONJUNTO, Y, O, LET, LAMBDA, MAP, REDUCE |
| **Fecha y Hora** | HOY, SIFECHA, DIAS.LAB, FIN.MES, DIASEM |
| **Base de Datos** | BDSUMA, BDCONTAR, BDPROMEDIO, BDMAX, BDMIN |
| **Matemáticas** | SUMA, REDONDEAR, ALEATORIO, SECUENCIA, MMULT, POTENCIA |
| **Financieras** | PAGO, VA, VF, TIR, VNA, SLN, DDB, NPER, TASA |
| **Estadísticas** | PROMEDIO, DESVEST.M, MEDIANA, MODA.UNO, PEARSON, TENDENCIA |
| **Información** | ESERROR, ESNUMERO, TIPO, CELDA, ESBLANCO |
| **Ingeniería** | CONVERTIR, DEC.A.BIN, COMPLEJO, IM.SUM |
| **Cubo** | CONJUNTOCUBO, VALORCUBO, MIEMBROCUBO |
| **Web** | SERVICIOWEB, XMLFILTRO, URLCODIF |

Cada función incluye nombre ES/EN, descripción, sintaxis y ejemplo. Botón "Insertar en Canvas".

### 4. Sibö ✦ Asistente Contable
Chat integrado que responde preguntas sobre:
- Cómo hacer facturas, asientos, balances
- Cálculo de IVA CR
- Uso de plantillas
- Conciliación bancaria
- Reportes D-151

Responde en el idioma seleccionado.

### 5. 10 Lenguas de Costa Rica
Cambio instantáneo en header: Español CR, English US, Bribri, Cabécar, Ngäbe/Guaymí, Boruca, Térraba, Maleku, Guaymí, Brunca. 
Traduce UI completa: botones, labels, modales, mensajes de Sibö, alertas.

### 6. Offline First + Guardado Local
- No necesita internet después de cargar
- `localStorage` guarda el canvas automáticamente
- Botón "Guardar" con feedback visual
- Botón "Compartir" usa Web Share API o copia link

## Stack Técnico

| Componente | Detalle |
| --- | --- |
| **Tipo** | SPA vanilla - 1 archivo HTML |
| **CSS** | Variables CSS, glassmorphism, backdrop-filter, gradientes |
| **Fonts** | Plus Jakarta Sans + JetBrains Mono vía Google Fonts |
| **JS** | Vanilla ES6, sin frameworks ni dependencias |
| **Almacenamiento** | localStorage para canvas y idioma |
| **i18n** | Objeto `translations` con 10 idiomas, mapa `i18nMap` dinámico |
| **Iconos** | SVG inline |
| **Responsive** | Grid + flex, breakpoint 640px |

## Instalación y Uso

1. **Descargar**: Clona el repo o descarga `index.html`
2. **Abrir**: Doble clic al archivo o súbelo a cualquier hosting estático. Funciona en móvil y desktop.
3. **Cambiar idioma**: Selector arriba derecha. Sibö y toda la UI cambian al instante.
4. **Crear asiento**: `Plantillas → Libro Diario` o `Tabla` en toolbar. Edita directo en las celdas.
5. **Emitir factura**: Botón `Factura` → llena emisor/receptor → agrega líneas → elige tasa IVA → `Enviar a Hacienda` o `XML`/`PDF`.
6. **Usar funciones**: `fx Funciones` → busca BUSCARV, SUMA, etc → `Insertar en Canvas`.
7. **Guardar trabajo**: `Guardar` en toolbar. Se restaura al recargar.

## Estructura del Proyecto
