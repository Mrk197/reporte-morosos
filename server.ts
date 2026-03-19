import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import { runScript, generateExcel, generateCSV, ScriptResult } from './new-clients-unpaid.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();
const PORT = process.env.PORT ? parseInt(process.env.PORT) : 3000;

// Historial en memoria (se reinicia con el servidor)
const history: Array<{
  timestamp: string;
  formato: string;
  perfil: string;
  categoria: string;
  filtroFecha: string;
  resumen: unknown;
  totales: { retirarModem: number; suspendidosSinPago: number; sinFacturaAun: number };
}> = [];

// Último resultado cacheado (para descarga rápida)
let lastResult: ScriptResult | null = null;

app.use(express.json());
app.use(express.static(__dirname));

// GET: Página principal
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// GET: Últimos resultados
app.get('/api/results', (req, res) => {
  res.json(lastResult);
});

// GET: Histórico
app.get('/api/history', (req, res) => {
  res.json(history);
});

// POST: Ejecutar script
app.post('/api/run-script', async (req, res) => {
  const {
    format = 'json',
    pages = '20',
    perfil = '2026',
    categoria = 'retirar',
    diasInstalacion = '',
    desdeFecha = '',
    hastaFecha = '',
  } = req.body;

  res.writeHead(200, {
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    'Connection': 'keep-alive',
  });

  const sendEvent = (data: Record<string, unknown>) => {
    res.write(`data: ${JSON.stringify(data)}\n\n`);
  };

  sendEvent({ status: 'iniciando', mensaje: 'Iniciando script...' });

  try {
    const result = await runScript({
      perfil: String(perfil),
      pages: parseInt(String(pages), 10),
      categoria: String(categoria),
      diasInstalacion: String(diasInstalacion),
      desdeFecha: String(desdeFecha),
      hastaFecha: String(hastaFecha),
      onProgress: (msg) => sendEvent({ status: 'procesando', linea: msg }),
    });

    lastResult = result;

    // Guardar en histórico
    let filtroFecha = 'Sin filtro';
    if (diasInstalacion) {
      filtroFecha = `Últimos ${diasInstalacion} días`;
    } else if (desdeFecha || hastaFecha) {
      filtroFecha = `${desdeFecha || 'inicio'} a ${hastaFecha || 'hoy'}`;
    }

    history.push({
      timestamp: new Date().toISOString(),
      formato: String(format),
      perfil: String(perfil),
      categoria: String(categoria),
      filtroFecha,
      resumen: result.resumen,
      totales: {
        retirarModem: result.retirarModem.length,
        suspendidosSinPago: result.suspendidosSinPago.length,
        sinFacturaAun: result.sinFacturaAun.length,
      },
    });

    sendEvent({
      status: 'completo',
      resultado: result,
      mensaje: 'Script ejecutado exitosamente',
    });
  } catch (err) {
    sendEvent({
      status: 'error',
      mensaje: `Error: ${err instanceof Error ? err.message : 'Error desconocido'}`,
    });
  }

  res.end();
});

// GET: Descargar Excel/CSV (genera en memoria, sin escribir a disco)
app.get('/api/download/:format', async (req, res) => {
  const { format } = req.params;
  const {
    pages = '20',
    perfil = '2026',
    categoria = 'retirar',
    diasInstalacion = '',
    desdeFecha = '',
    hastaFecha = '',
  } = req.query;

  try {
    const result = await runScript({
      perfil: String(perfil),
      pages: parseInt(String(pages), 10),
      categoria: String(categoria),
      diasInstalacion: String(diasInstalacion),
      desdeFecha: String(desdeFecha),
      hastaFecha: String(hastaFecha),
    });

    const cat = String(categoria);
    const lista = cat === 'sinFactura' ? result.sinFacturaAun
      : cat === 'todos' ? [...result.retirarModem, ...result.suspendidosSinPago, ...result.sinFacturaAun]
      : result.retirarModem;

    const dateStr = new Date().toISOString().split('T')[0];
    const basename = cat === 'sinFactura' ? 'sin-factura-aun'
      : cat === 'todos' ? 'todos-sin-pagar'
      : 'retirar-modem';

    if (format === 'excel') {
      const buffer = await generateExcel(lista, cat === 'todos', {
        retirar: result.retirarModem.length,
        suspendidos: result.suspendidosSinPago.length,
        sinFactura: result.sinFacturaAun.length,
      });
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', `attachment; filename="${basename}-${dateStr}.xlsx"`);
      res.send(buffer);
    } else if (format === 'csv') {
      const csv = generateCSV(lista, cat === 'todos');
      res.setHeader('Content-Type', 'text/csv; charset=utf-8');
      res.setHeader('Content-Disposition', `attachment; filename="${basename}-${dateStr}.csv"`);
      res.send(csv);
    } else {
      res.json(result);
    }
  } catch (err) {
    res.status(500).json({ error: `Error: ${err instanceof Error ? err.message : 'Error desconocido'}` });
  }
});

// Solo escuchar en local; en Vercel se exporta el app
if (!process.env.VERCEL) {
  app.listen(PORT, () => {
    console.log(`\n✅ Servidor iniciado en http://localhost:${PORT}`);
    console.log(`📊 Abre tu navegador en http://localhost:${PORT}\n`);
  });
}

export default app;
