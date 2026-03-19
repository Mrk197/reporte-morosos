/**
 * Clientes nuevos sin primer pago
 *
 * Exporta runScript(), generateExcel(), generateCSV() para uso como módulo.
 * También funciona como CLI: tsx new-clients-unpaid.ts --format excel
 */

import { mikroWispPostRaw, extractFacturas } from './mikrowisp.js';
import ExcelJS from 'exceljs';
import { fileURLToPath } from 'url';

// ============================================================
// Types
// ============================================================

export interface ClienteResult {
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

export interface ScriptParams {
  perfil?: string;
  pages?: number;
  categoria?: string;
  diasInstalacion?: string;
  desdeFecha?: string;
  hastaFecha?: string;
  onProgress?: (msg: string) => void;
}

export interface ScriptResult {
  resumen: {
    totalClientes: string;
    paginasRevisadas: string;
    clientesRevisados: number;
    retirarModem: number;
    suspendidosSinPago: number;
    sinFacturaAun: number;
  };
  retirarModem: ClienteResult[];
  suspendidosSinPago: ClienteResult[];
  sinFacturaAun: ClienteResult[];
  errores?: Array<{ id: number; error: string }>;
  consultadoEn: string;
}

// ============================================================
// Main exported function
// ============================================================

export async function runScript(params: ScriptParams = {}): Promise<ScriptResult> {
  const {
    perfil: perfilBuscado = '2026',
    pages: pagesToCheck = 20,
    diasInstalacion = '',
    desdeFecha = '',
    hastaFecha = '',
    onProgress = () => {},
  } = params;

  const shouldInclude = (fechaInstalado: string) =>
    checkDateFilter(fechaInstalado, { diasInstalacion, desdeFecha, hastaFecha });

  onProgress('=== Clientes nuevos sin primer pago ===');
  onProgress(`Perfil: "${perfilBuscado}" | Páginas: ${pagesToCheck}`);

  if (diasInstalacion) {
    onProgress(`Filtro: Últimos ${diasInstalacion} días desde instalación`);
  } else if (desdeFecha || hastaFecha) {
    onProgress(`Filtro: Instalación entre ${desdeFecha || 'inicio'} y ${hastaFecha || 'hoy'}`);
  }

  // PASO 1: Obtener clientes de las últimas páginas
  const totalPages = await findLastPage(onProgress);
  const startPage = Math.max(1, totalPages - pagesToCheck + 1);

  onProgress(`Páginas ${startPage}-${totalPages}, buscando perfil "${perfilBuscado}"`);

  const clientIds: Array<{ id: number; nombre: string; estado: string }> = [];

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
    onProgress(`Obteniendo IDs... ${clientIds.length} clientes`);
  }

  // Solo SUSPENDIDOS
  const suspendidos = clientIds.filter((c) => c.estado === 'SUSPENDIDO');
  onProgress(`${clientIds.length} IDs obtenidos | ${suspendidos.length} SUSPENDIDOS`);

  // PASO 2: Consultar detalles y filtrar
  const retirarModem: ClienteResult[] = [];
  const suspendidosSinPago: ClienteResult[] = [];
  const sinFacturaAun: ClienteResult[] = [];
  const errores: Array<{ id: number; error: string }> = [];
  const BATCH = 20;
  let processed = 0;

  for (let i = 0; i < suspendidos.length; i += BATCH) {
    const batch = suspendidos.slice(i, i + BATCH);

    const batchResults = await Promise.all(
      batch.map((client) => processClient(client.id, perfilBuscado, shouldInclude))
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
    onProgress(`Procesando... ${processed}/${suspendidos.length} (retirar: ${retirarModem.length}, sinFactura: ${sinFacturaAun.length})`);

    if (i + BATCH < suspendidos.length) await sleep(50);
  }

  // Ordenar: más antiguos primero
  const byDate = (a: ClienteResult, b: ClienteResult) =>
    (a.servicio.instalado || '9999').localeCompare(b.servicio.instalado || '9999');
  retirarModem.sort(byDate);
  suspendidosSinPago.sort(byDate);
  sinFacturaAun.sort(byDate);

  onProgress('--- RESUMEN ---');
  onProgress(`Retirar modem: ${retirarModem.length}`);
  onProgress(`Suspendidos sin pago: ${suspendidosSinPago.length}`);
  onProgress(`Sin factura aún: ${sinFacturaAun.length}`);
  if (errores.length > 0) onProgress(`Errores: ${errores.length}`);

  return {
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
  };
}

// ============================================================
// Process a single client
// ============================================================

async function processClient(
  id: number,
  perfilBuscado: string,
  shouldInclude: (fechaInstalado: string) => boolean
): Promise<{ data?: ClienteResult; id?: number; error?: string } | null> {
  try {
    const res = await mikroWispPostRaw('GetClientsDetails', { idcliente: id });
    const raw = res.data as Record<string, unknown>;
    if (raw.estado !== 'exito' || !raw.datos) return null;

    const datos = (raw.datos as Array<Record<string, unknown>>)[0];
    const servicios = (datos.servicios || []) as Array<Record<string, unknown>>;

    const svc = servicios.find((s) =>
      String(s.perfil || '').toLowerCase().includes(perfilBuscado.toLowerCase())
    );
    if (!svc) return null;

    const fechaInstalado = String(svc.instalado || '');
    if (!shouldInclude(fechaInstalado)) return null;

    const facturacion = datos.facturacion as Record<string, unknown> | undefined;
    const noPagadas = Number(facturacion?.facturas_nopagadas || 0);
    const totalDeuda = String(facturacion?.total_facturas || '0.00');

    const [paidResult, pendResult] = await Promise.all([
      mikroWispPostRaw('GetInvoices', { idcliente: id, estado: 0 }).catch(() => null),
      noPagadas > 0
        ? mikroWispPostRaw('GetInvoices', { idcliente: id, estado: 1 }).catch(() => null)
        : Promise.resolve(null),
    ]);

    if (paidResult) {
      const paidData = paidResult.data as Record<string, unknown>;
      if (paidData.estado === 'exito' && extractFacturas(paidData).length > 0) return null;
    }

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

async function findLastPage(onProgress: (msg: string) => void): Promise<number> {
  onProgress('Buscando última página...');
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
  onProgress(`Última página: ${last}`);
  return last;
}

interface FilterParams {
  diasInstalacion: string;
  desdeFecha: string;
  hastaFecha: string;
}

function checkDateFilter(dateStr: string, params: FilterParams): boolean {
  if (params.diasInstalacion) {
    const dias = parseInt(params.diasInstalacion, 10);
    if (isNaN(dias) || dias < 0) return true;
    return isDateInRange(dateStr, 0, dias);
  }
  if (params.desdeFecha || params.hastaFecha) {
    const desde = params.desdeFecha ? parseDateParam(params.desdeFecha) : null;
    const hasta = params.hastaFecha ? parseDateParam(params.hastaFecha) : null;
    return isDateBetween(dateStr, desde, hasta);
  }
  return true;
}

function isDateInRange(dateStr: string, minDays: number | null, maxDays: number | null): boolean {
  if (!dateStr) return true;
  try {
    const installDate = new Date(dateStr);
    const today = new Date();
    const daysAgo = Math.floor((today.getTime() - installDate.getTime()) / (1000 * 60 * 60 * 24));
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

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

// ============================================================
// Excel / CSV generators (exported para uso desde server.ts)
// ============================================================

export async function generateExcel(
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
    { key: 'conexion', header: 'Conexión', width: 11 },
    { key: 'deuda', header: 'Deuda', width: 12 },
    { key: 'vencimiento', header: 'Vencimiento', width: 14 },
    { key: 'nodo', header: 'Nodo', width: 7 },
    { key: 'pppuser', header: 'PPP User', width: 24 },
    { key: 'coordenadas', header: 'Coordenadas', width: 26 },
    ...(includeCategoria ? [{ key: 'categoria', header: 'Categoría', width: 16 }] : []),
  ];

  headers.forEach((h, i) => {
    ws.getColumn(i + 1).width = h.width;
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
    const row = ws.getRow(idx + 5);
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
      c.servicio.statusUser,
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

    const bgColor = idx % 2 === 0 ? 'FFFFFFFF' : 'FFF9FAFB';
    values.forEach((_, i) => {
      row.getCell(i + 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
    });

    row.getCell(9).numFmt = '$#,##0.00';
    row.getCell(13).numFmt = '$#,##0.00';

    const conexionCell = row.getCell(12);
    if (c.servicio.statusUser === 'OFFLINE') {
      conexionCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFDC2626' } };
    } else if (c.servicio.statusUser === 'ONLINE') {
      conexionCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FF16A34A' } };
    }

    const deudaCell = row.getCell(13);
    if (Number(c.facturacion.totalDeuda) > 0) {
      deudaCell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFDC2626' } };
    }

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

  const lastCol = includeCategoria ? 18 : 17;
  ws.autoFilter = {
    from: { row: 4, column: 1 },
    to: { row: 4 + clientes.length, column: lastCol },
  };

  const buffer = await wb.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

export function generateCSV(clientes: ClienteResult[], includeCategoria: boolean): string {
  const BOM = '\uFEFF';
  const headers = [
    'ID', 'Nombre', 'Estado', 'Telefono', 'Celular', 'Correo',
    'Direccion', 'Coordenadas', 'Plan', 'Costo', 'Fecha Instalacion',
    'Dias Instalado', 'Conexion', 'Nodo', 'PPP User',
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
      String(c.servicio.diasDesdeInstalacion), c.servicio.statusUser,
      String(c.servicio.nodo), c.servicio.pppuser,
      String(c.facturacion.facturasPendientes), '$' + c.facturacion.totalDeuda,
      String(factura1?.vencimiento || 'N/A'),
      ...(includeCategoria ? [cat] : []),
    ].join(',');
  });

  return BOM + headers.join(',') + '\n' + rows.join('\n');
}

// ============================================================
// CLI Entry Point (solo para uso local con tsx)
// ============================================================

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const argv = process.argv.slice(2);
  const getArg = (name: string, def: string) => {
    const idx = argv.indexOf(`--${name}`);
    return idx === -1 ? def : argv[idx + 1] || def;
  };

  const format = getArg('format', 'json');
  const categoria = getArg('categoria', 'retirar');

  runScript({
    perfil: getArg('perfil', '2026'),
    pages: parseInt(getArg('pages', '20'), 10),
    categoria,
    diasInstalacion: getArg('dias-instalacion', ''),
    desdeFecha: getArg('desde-fecha', ''),
    hastaFecha: getArg('hasta-fecha', ''),
    onProgress: (msg) => process.stderr.write('\r' + msg),
  }).then(async (result) => {
    process.stderr.write('\n');

    if (format === 'excel' || format === 'csv') {
      const { default: fs } = await import('fs');
      const { default: path } = await import('path');
      const lista = categoria === 'sinFactura' ? result.sinFacturaAun
        : categoria === 'todos' ? [...result.retirarModem, ...result.suspendidosSinPago, ...result.sinFacturaAun]
        : result.retirarModem;
      const basename = categoria === 'sinFactura' ? 'sin-factura-aun'
        : categoria === 'todos' ? 'todos-sin-pagar' : 'retirar-modem';
      const timestamp = new Date().toISOString().replace(/[:.]/g, '').substring(0, 15);

      if (format === 'excel') {
        const buffer = await generateExcel(lista, categoria === 'todos', {
          retirar: result.retirarModem.length,
          suspendidos: result.suspendidosSinPago.length,
          sinFactura: result.sinFacturaAun.length,
        });
        const outPath = path.resolve(`${basename}-${timestamp}.xlsx`);
        fs.writeFileSync(outPath, buffer);
        process.stderr.write(`Excel generado: ${outPath}\n`);
      } else {
        const csv = generateCSV(lista, categoria === 'todos');
        const outPath = path.resolve(`${basename}-${timestamp}.csv`);
        fs.writeFileSync(outPath, csv, 'utf-8');
        process.stderr.write(`CSV generado: ${outPath}\n`);
      }
    }

    console.log(JSON.stringify(result, null, 2));
  }).catch((err) => {
    console.error('Error fatal:', err);
    process.exit(1);
  });
}
