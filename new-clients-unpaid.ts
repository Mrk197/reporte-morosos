/**
 * Clientes nuevos sin primer pago - Script standalone
 *
 * Identifica clientes de planes "2026" que no pagaron su primer servicio.
 * Genera JSON, CSV o Excel con los resultados.
 *
 * Uso:
 *   npx tsx new-clients-unpaid.ts                     # JSON en consola
 *   npx tsx new-clients-unpaid.ts --format excel      # Genera .xlsx
 *   npx tsx new-clients-unpaid.ts --format csv        # Genera .csv
 *   npx tsx new-clients-unpaid.ts --test --id 23844   # Consultar 1 cliente
 *   npx tsx new-clients-unpaid.ts --pages 30          # Últimas 30 páginas
 *   npx tsx new-clients-unpaid.ts --perfil "2026"     # Filtro de perfil
 *   npx tsx new-clients-unpaid.ts --categoria todos   # retirar|sinFactura|todos
 *
 * Filtros por fecha de instalación:
 *   npx tsx new-clients-unpaid.ts --dias-instalacion 30     # Últimos 30 días
 *   npx tsx new-clients-unpaid.ts --dias-instalacion 14     # Últimas 2 semanas
 *   npx tsx new-clients-unpaid.ts --desde-fecha 2026-01-01 --hasta-fecha 2026-02-01
 *   npx tsx new-clients-unpaid.ts --desde-fecha 2026-01-15  # Desde fecha hasta hoy
 *   npx tsx new-clients-unpaid.ts --hasta-fecha 2026-02-01  # Desde inicio hasta fecha
 */

import { mikroWispPostRaw, extractFacturas } from './mikrowisp.js';
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';

// ============================================================
// CLI Args
// ============================================================

const args = process.argv.slice(2);
function getArg(name: string, def: string): string {
  const idx = args.indexOf(`--${name}`);
  if (idx === -1) return def;
  return args[idx + 1] || def;
}
function hasFlag(name: string): boolean {
  return args.includes(`--${name}`);
}

const perfilBuscado = getArg('perfil', '2026');
const pagesToCheck = parseInt(getArg('pages', '20'), 10);
const format = getArg('format', 'json');
const categoriaCSV = getArg('categoria', 'retirar');
const testMode = hasFlag('test');
const testId = getArg('id', '');

// Filtros de fecha de instalación
const diasInstalacion = getArg('dias-instalacion', '');
const desdeFecha = getArg('desde-fecha', '');
const hastaFecha = getArg('hasta-fecha', '');

// ============================================================
// Types
// ============================================================

interface ClienteResult {
  id: number;
  nombre: string;
  estado: string;
  correo: string;
  movil: string;
  telefono: string;
  direccion: string;
  coordenada: string;
  servicio: {
    perfil: string;
    costo: string;
    instalado: string;
    diasDesdeInstalacion: number;
    statusUser: string;
    nodo: number;
    pppuser: string;
  };
  facturacion: {
    facturasPendientes: number;
    totalDeuda: string;
    facturasPagadas: number;
    detalleFacturasPendientes: Array<Record<string, unknown>>;
  };
}

// ============================================================
// Main
// ============================================================

async function main() {
  console.log('=== Clientes nuevos sin primer pago ===\n');

  if (testMode && testId) {
    const result = await queryOneClient(Number(testId));
    console.log(JSON.stringify(result, null, 2));
    return;
  }

  // === Mostrar parámetros activos ===
  console.log('--- PARÁMETROS ---');
  console.log(`  Perfil: "${perfilBuscado}"`);
  console.log(`  Páginas: ${pagesToCheck}`);
  
  if (diasInstalacion) {
    console.log(`  Filtro: Últimos ${diasInstalacion} días desde instalación`);
  } else if (desdeFecha || hastaFecha) {
    const desde = desdeFecha || 'inicio';
    const hasta = hastaFecha || 'hoy';
    console.log(`  Filtro: Instalación entre ${desde} y ${hasta}`);
  } else {
    console.log(`  Filtro: Sin restricción de fecha`);
  }
  
  console.log(`  Categoría: ${categoriaCSV}`);
  console.log(`  Formato: ${format}`);
  console.log('');

  // === PASO 1: Obtener clientes de las últimas páginas ===
  const totalPages = await findLastPage();
  const startPage = Math.max(1, totalPages - pagesToCheck + 1);

  console.log(`Páginas ${startPage}-${totalPages}, buscando perfil "${perfilBuscado}"`);

  const clientIds: Array<{ id: number; nombre: string; estado: string }> = [];

  // Obtener IDs en lotes de 5 páginas en paralelo
  for (let p = startPage; p <= totalPages; p += 5) {
    const pages = [];
    for (let pp = p; pp < p + 5 && pp <= totalPages; pp++) pages.push(pp);

    const results = await Promise.all(
      pages.map(async (pg) => {
        const res = await mikroWispPostRaw('GetAllClients', { pagina: pg, limit: 100 });
        const data = res.data as Record<string, unknown>;
        return ((data.clientes || []) as Array<Record<string, unknown>>).map((c) => ({
          id: Number(c.id),
          nombre: String(c.nombre || ''),
          estado: String(c.estado || ''),
        }));
      })
    );
    for (const r of results) clientIds.push(...r);
    process.stdout.write(`\r  Obteniendo IDs... ${clientIds.length} clientes`);
  }
  console.log(`\n  ${clientIds.length} IDs obtenidos`);

  // === FILTRO: Solo clientes SUSPENDIDOS ===
  const suspendidos = clientIds.filter((c) => c.estado === 'SUSPENDIDO');
  console.log(`  Filtrados: ${suspendidos.length} clientes SUSPENDIDOS (descartados ${clientIds.length - suspendidos.length} no suspendidos)`);

  // === PASO 2: Consultar detalles y filtrar ===
  const retirarModem: ClienteResult[] = [];
  const suspendidosSinPago: ClienteResult[] = [];
  const sinFacturaAun: ClienteResult[] = [];
  const errores: Array<{ id: number; error: string }> = [];
  const BATCH = 20;
  let processed = 0;

  for (let i = 0; i < suspendidos.length; i += BATCH) {
    const batch = suspendidos.slice(i, i + BATCH);

    const batchResults = await Promise.all(
      batch.map((client) => processClient(client.id, perfilBuscado))
    );

    for (const r of batchResults) {
      if (!r) continue;
      if (r.error) { errores.push({ id: r.id!, error: r.error }); continue; }
      if (!r.data) continue;

      const c = r.data;
      if (c.facturacion.facturasPendientes > 0) {
        retirarModem.push(c);
      } else if (c.estado === 'SUSPENDIDO') {
        suspendidosSinPago.push(c);
      } else {
        sinFacturaAun.push(c);
      }
    }

    processed += batch.length;
    process.stdout.write(`\r  Procesando... ${processed}/${suspendidos.length} (retirar: ${retirarModem.length}, sinFactura: ${sinFacturaAun.length})`);

    if (i + BATCH < suspendidos.length) await sleep(50);
  }

  console.log('\n');

  // Ordenar: más antiguos primero
  const byDate = (a: ClienteResult, b: ClienteResult) =>
    (a.servicio.instalado || '9999').localeCompare(b.servicio.instalado || '9999');
  retirarModem.sort(byDate);
  suspendidosSinPago.sort(byDate);
  sinFacturaAun.sort(byDate);

  // === Resumen ===
  console.log('--- RESUMEN ---');
  console.log(`  Total clientes: ~${totalPages * 100}`);
  console.log(`  Páginas revisadas: ${startPage}-${totalPages}`);
  console.log(`  Clientes en páginas: ${clientIds.length}`);
  console.log(`  Suspendidos encontrados: ${suspendidos.length}`);
  console.log(`  Retirar modem: ${retirarModem.length}`);
  console.log(`  Suspendidos sin pago: ${suspendidosSinPago.length}`);
  console.log(`  Sin factura aún: ${sinFacturaAun.length}`);
  if (errores.length > 0) console.log(`  Errores: ${errores.length}`);
  console.log('');

  // === Exportar ===
  let lista: ClienteResult[];
  let filename: string;
  switch (categoriaCSV) {
    case 'sinFactura': lista = sinFacturaAun; filename = 'sin-factura-aun'; break;
    case 'todos': lista = [...retirarModem, ...suspendidosSinPago, ...sinFacturaAun]; filename = 'todos-sin-pagar'; break;
    default: lista = retirarModem; filename = 'retirar-modem'; break;
  }
  
  if (format === 'excel') {
    const buffer = await generateExcel(lista, true, {
      retirar: retirarModem.length,
      suspendidos: suspendidosSinPago.length,
      sinFactura: sinFacturaAun.length,
    });
    const outPath = path.resolve(`${filename}.xlsx`);
    fs.writeFileSync(outPath, buffer);
    console.log(`Excel generado: ${outPath}`);
  } else if (format === 'csv') {
    const csv = generateCSV(lista, categoriaCSV === 'todos');
    const outPath = path.resolve(`${filename}.csv`);
    fs.writeFileSync(outPath, csv, 'utf-8');
    console.log(`CSV generado: ${outPath}`);
  }

  // SIEMPRE generar JSON en stdout para que el servidor pueda parsear
  console.log(JSON.stringify({
    resumen: {
      totalClientes: `~${totalPages * 100}`,
      paginasRevisadas: `${startPage}-${totalPages}`,
      clientesRevisados: clientIds.length,
      retirarModem: retirarModem.length,
      suspendidosSinPago: suspendidosSinPago.length,
      sinFacturaAun: sinFacturaAun.length,
    },
    retirarModem,
    suspendidosSinPago,
    sinFacturaAun,
    ...(errores.length > 0 && { errores }),
    consultadoEn: new Date().toISOString(),
  }, null, 2));
}

// ============================================================
// Process a single client
// ============================================================

async function processClient(
  id: number,
  perfilBuscado: string
): Promise<{ data?: ClienteResult; id?: number; error?: string } | null> {
  try {
    const res = await mikroWispPostRaw('GetClientsDetails', { idcliente: id });
    const raw = res.data as Record<string, unknown>;
    if (raw.estado !== 'exito' || !raw.datos) return null;

    const datos = (raw.datos as Array<Record<string, unknown>>)[0];
    const servicios = (datos.servicios || []) as Array<Record<string, unknown>>;

    // Buscar servicio del plan 2026
    const svc = servicios.find((s) =>
      String(s.perfil || '').toLowerCase().includes(perfilBuscado.toLowerCase())
    );
    if (!svc) return null;

    // Obtener fecha de instalación
    const fechaInstalado = String(svc.instalado || '');
    
    // Filtrar por fecha de instalación
    if (!shouldIncludeClient(fechaInstalado)) return null;

    // Verificar facturas (pagadas + pendientes en paralelo)
    const facturacion = datos.facturacion as Record<string, unknown> | undefined;
    const noPagadas = Number(facturacion?.facturas_nopagadas || 0);
    const totalDeuda = String(facturacion?.total_facturas || '0.00');

    const [paidResult, pendResult] = await Promise.all([
      mikroWispPostRaw('GetInvoices', { idcliente: id, estado: 0 }).catch(() => null),
      noPagadas > 0
        ? mikroWispPostRaw('GetInvoices', { idcliente: id, estado: 1 }).catch(() => null)
        : Promise.resolve(null),
    ]);

    // Si tiene pagos, descartar
    if (paidResult) {
      const paidData = paidResult.data as Record<string, unknown>;
      if (paidData.estado === 'exito' && extractFacturas(paidData).length > 0) return null;
    }

    // Detalle de facturas pendientes
    let facturasPendientesDetalle: Array<Record<string, unknown>> = [];
    if (pendResult) {
      const pendData = pendResult.data as Record<string, unknown>;
      if (pendData.estado === 'exito') {
        facturasPendientesDetalle = extractFacturas(pendData);
      }
    }

    const hoy = new Date();
    const diasDesdeInstalacion = fechaInstalado
      ? Math.floor((hoy.getTime() - new Date(fechaInstalado).getTime()) / (1000 * 60 * 60 * 24))
      : 0;

    return {
      data: {
        id: Number(datos.id),
        nombre: String(datos.nombre || ''),
        estado: String(datos.estado || ''),
        correo: String(datos.correo || ''),
        movil: String(datos.movil || ''),
        telefono: String(datos.telefono || ''),
        direccion: String(datos.direccion_principal || svc.direccion || ''),
        coordenada: String(datos.coordenada || svc.coordenadas || ''),
        servicio: {
          perfil: String(svc.perfil || ''),
          costo: String(svc.costo || ''),
          instalado: fechaInstalado,
          diasDesdeInstalacion,
          statusUser: String(svc.status_user || ''),
          nodo: Number(svc.nodo || 0),
          pppuser: String(svc.pppuser || ''),
        },
        facturacion: {
          facturasPendientes: noPagadas,
          totalDeuda,
          facturasPagadas: 0,
          detalleFacturasPendientes: facturasPendientesDetalle.map((f) => ({
            id: f.id || f.idfactura,
            total: f.total || f.monto,
            fecha: f.fecha || f.fecha_emision,
            vencimiento: f.vencimiento || f.fecha_vencimiento,
          })),
        },
      },
    };
  } catch (err) {
    return { id, error: err instanceof Error ? err.message : 'Error' };
  }
}

// ============================================================
// Helpers
// ============================================================

async function findLastPage(): Promise<number> {
  process.stdout.write('  Buscando última página...');
  let low = 1, high = 300, last = 1;
  while (low <= high) {
    const mid = Math.floor((low + high) / 2);
    try {
      const res = await mikroWispPostRaw('GetAllClients', { pagina: mid, limit: 100 });
      const data = res.data as Record<string, unknown>;
      if (((data.clientes || []) as Array<unknown>).length > 0) { last = mid; low = mid + 1; }
      else high = mid - 1;
    } catch { high = mid - 1; }
  }
  console.log(` ${last}`);
  return last;
}

// ============================================================
// Filter Functions - Fecha de Instalación
// ============================================================

function isDateInRange(dateStr: string, minDays: number | null, maxDays: number | null): boolean {
  if (!dateStr) return true; // Si no hay fecha, incluir
  
  try {
    const installDate = new Date(dateStr);
    const today = new Date();
    const daysAgo = Math.floor((today.getTime() - installDate.getTime()) / (1000 * 60 * 60 * 24));
    
    // Validar rango de días
    if (minDays !== null && daysAgo < minDays) return false;
    if (maxDays !== null && daysAgo > maxDays) return false;
    
    return true;
  } catch {
    return true;
  }
}

function isDateBetween(dateStr: string, desde: Date | null, hasta: Date | null): boolean {
  if (!dateStr) return true;
  
  try {
    const date = new Date(dateStr);
    if (desde && date < desde) return false;
    if (hasta && date > hasta) return false;
    return true;
  } catch {
    return true;
  }
}

function parseDateParam(dateStr: string): Date | null {
  if (!dateStr) return null;
  try {
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) return null;
    return date;
  } catch {
    return null;
  }
}

function shouldIncludeClient(fechaInstalado: string): boolean {
  // Opción 1: Filtro por días desde instalación
  if (diasInstalacion) {
    const dias = parseInt(diasInstalacion, 10);
    if (isNaN(dias) || dias < 0) {
      console.error(`⚠️  --dias-instalacion debe ser un número positivo`);
      return true;
    }
    // Incluir clientes instalados en los últimos N días
    // diasDesdeInstalacion va de 0 (hoy) hacia atrás
    return isDateInRange(fechaInstalado, 0, dias);
  }

  // Opción 2: Filtro por rango de fechas
  if (desdeFecha || hastaFecha) {
    const desde = desdeFecha ? parseDateParam(desdeFecha) : null;
    const hasta = hastaFecha ? parseDateParam(hastaFecha) : null;

    if (desdeFecha && !desde) {
      console.error(`⚠️  --desde-fecha inválido. Usa formato: YYYY-MM-DD (ej: 2026-01-15)`);
    }
    if (hastaFecha && !hasta) {
      console.error(`⚠️  --hasta-fecha inválido. Usa formato: YYYY-MM-DD (ej: 2026-02-01)`);
    }

    return isDateBetween(fechaInstalado, desde, hasta);
  }

  // Sin filtro de fecha
  return true;
}

async function queryOneClient(id: number) {
  console.log(`Consultando cliente ${id}...`);
  const res = await mikroWispPostRaw('GetClientsDetails', { idcliente: id });
  let pend = null, paid = null;
  try {
    const [p1, p2] = await Promise.all([
      mikroWispPostRaw('GetInvoices', { idcliente: id, estado: 1 }),
      mikroWispPostRaw('GetInvoices', { idcliente: id, estado: 0 }),
    ]);
    pend = p1.data; paid = p2.data;
  } catch { /* skip */ }
  return { clienteDetalle: res.data, facturasPendientes: pend, facturasPagadas: paid };
}

async function generateExcel(
  clientes: ClienteResult[],
  includeCategoria: boolean,
  resumen: { retirar: number; suspendidos: number; sinFactura: number }
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'DIGY - MikroWisp Scripts';
  wb.created = new Date();

  const ws = wb.addWorksheet('Retirar Modem', {
    views: [{ state: 'frozen', ySplit: 4 }],
  });

  // --- Título ---
  ws.mergeCells('A1:R1');
  const titleCell = ws.getCell('A1');
  titleCell.value = 'REPORTE DE MODEMS A RETIRAR - CLIENTES SIN PRIMER PAGO';
  titleCell.font = { name: 'Calibri', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A5F' } };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 35;

  // --- Resumen ---
  ws.mergeCells('A2:F2');
  const fechaHoy = new Date().toLocaleDateString('es-MX', { day: '2-digit', month: 'long', year: 'numeric' });
  ws.getCell('A2').value = `Fecha: ${fechaHoy}`;
  ws.getCell('A2').font = { name: 'Calibri', size: 10, italic: true, color: { argb: 'FF555555' } };

  ws.getCell('G2').value = `Retirar: ${resumen.retirar}`;
  ws.getCell('G2').font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFDC2626' } };
  ws.getCell('I2').value = `Sin factura aún: ${resumen.sinFactura}`;
  ws.getCell('I2').font = { name: 'Calibri', size: 10, color: { argb: 'FF777777' } };
  ws.getCell('K2').value = `Total: ${clientes.length}`;
  ws.getCell('K2').font = { name: 'Calibri', size: 10, bold: true };
  ws.getRow(2).height = 20;

  // --- Fila separadora ---
  ws.getRow(3).height = 5;

  // --- Headers ---
  const headers = [
    { key: 'num', header: '#', width: 5 },
    { key: 'id', header: 'ID Cliente', width: 12 },
    { key: 'nombre', header: 'Nombre', width: 32 },
    { key: 'telefono', header: 'Teléfono', width: 14 },
    { key: 'celular', header: 'Celular', width: 14 },
    { key: 'correo', header: 'Correo', width: 28 },
    { key: 'direccion', header: 'Dirección', width: 50 },
    { key: 'plan', header: 'Plan', width: 30 },
    { key: 'costo', header: 'Costo', width: 10 },
    { key: 'instalacion', header: 'F. Instalación', width: 14 },
    { key: 'dias', header: 'Días', width: 7 },
    { key: 'deuda', header: 'Deuda', width: 12 },
    { key: 'vencimiento', header: 'Vencimiento', width: 14 },
    { key: 'nodo', header: 'Nodo', width: 7 },
    { key: 'pppuser', header: 'PPP User', width: 24 },
    { key: 'coordenadas', header: 'Coordenadas', width: 26 },
    ...(includeCategoria ? [{ key: 'categoria', header: 'Categoría', width: 16 }] : []),
  ];

  headers.forEach((h, i) => {
    const col = ws.getColumn(i + 1);
    col.width = h.width;
  });

  const headerRow = ws.getRow(4);
  headers.forEach((h, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = h.header;
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF1E40AF' } },
      bottom: { style: 'thin', color: { argb: 'FF1E40AF' } },
      left: { style: 'thin', color: { argb: 'FF1E40AF' } },
      right: { style: 'thin', color: { argb: 'FF1E40AF' } },
    };
  });
  headerRow.height = 25;

  // --- Data rows ---
  clientes.forEach((c, idx) => {
    const rowNum = idx + 5;
    const row = ws.getRow(rowNum);
    const factura1 = c.facturacion.detalleFacturasPendientes?.[0];
    const cat = c.facturacion.facturasPendientes > 0
      ? 'RETIRAR MODEM'
      : c.estado === 'SUSPENDIDO' ? 'SUSPENDIDO' : 'SIN FACTURA';

    const values: (string | number)[] = [
      idx + 1,
      c.id,
      c.nombre,
      c.telefono,
      c.movil,
      c.correo,
      c.direccion,
      c.servicio.perfil,
      Number(c.servicio.costo),
      c.servicio.instalado,
      c.servicio.diasDesdeInstalacion,
      Number(c.facturacion.totalDeuda),
      String(factura1?.vencimiento || 'N/A'),
      c.servicio.nodo,
      c.servicio.pppuser,
      c.coordenada,
      ...(includeCategoria ? [cat] : []),
    ];

    values.forEach((v, i) => {
      const cell = row.getCell(i + 1);
      cell.value = v;
      cell.font = { name: 'Calibri', size: 10 };
      cell.alignment = { vertical: 'middle', wrapText: i === 6 };
      cell.border = {
        bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } },
        left: { style: 'thin', color: { argb: 'FFE5E7EB' } },
        right: { style: 'thin', color: { argb: 'FFE5E7EB' } },
      };
    });

    // Zebra striping
    const bgColor = idx % 2 === 0 ? 'FFFFFFFF' : 'FFF9FAFB';
    values.forEach((_, i) => {
      row.getCell(i + 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
    });

    // Formato moneda
    row.getCell(9).numFmt = '$#,##0.00';
    row.getCell(12).numFmt = '$#,##0.00';

    // Color de deuda
    const deudaCell = row.getCell(12);
    if (Number(c.facturacion.totalDeuda) > 0) {
      deudaCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFDC2626' } };
    }

    // Color de categoría
    if (includeCategoria) {
      const catCell = row.getCell(values.length);
      if (cat === 'RETIRAR MODEM') {
        catCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        catCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDC2626' } };
      } else if (cat === 'SUSPENDIDO') {
        catCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        catCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF59E0B' } };
      }
    }

    row.height = 22;
  });

  // --- Auto filtro ---
  const lastCol = includeCategoria ? 17 : 16;
  ws.autoFilter = {
    from: { row: 4, column: 1 },
    to: { row: 4 + clientes.length, column: lastCol },
  };

  const buffer = await wb.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

function generateCSV(clientes: ClienteResult[], includeCategoria: boolean): string {
  const BOM = '\uFEFF';
  const headers = [
    'ID', 'Nombre', 'Estado', 'Telefono', 'Celular', 'Correo',
    'Direccion', 'Coordenadas', 'Plan', 'Costo', 'Fecha Instalacion',
    'Dias Instalado', 'Nodo', 'PPP User',
    'Facturas Pendientes', 'Deuda Total', 'Vencimiento',
    ...(includeCategoria ? ['Categoria'] : []),
  ];

  const esc = (v: string) =>
    v.includes(',') || v.includes('"') || v.includes('\n')
      ? '"' + v.replace(/"/g, '""') + '"'
      : v;

  const rows = clientes.map((c) => {
    const factura1 = c.facturacion.detalleFacturasPendientes?.[0];
    const cat = c.facturacion.facturasPendientes > 0
      ? 'RETIRAR MODEM'
      : c.estado === 'SUSPENDIDO' ? 'SUSPENDIDO' : 'SIN FACTURA';

    return [
      String(c.id), esc(c.nombre), c.estado, c.telefono, c.movil, c.correo,
      esc(c.direccion), c.coordenada, esc(c.servicio.perfil),
      '$' + c.servicio.costo, c.servicio.instalado,
      String(c.servicio.diasDesdeInstalacion),
      String(c.servicio.nodo), c.servicio.pppuser,
      String(c.facturacion.facturasPendientes), '$' + c.facturacion.totalDeuda,
      String(factura1?.vencimiento || 'N/A'),
      ...(includeCategoria ? [cat] : []),
    ].join(',');
  });

  return BOM + headers.join(',') + '\n' + rows.join('\n');
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

// Run
main().catch((err) => {
  console.error('Error fatal:', err);
  process.exit(1);
});
