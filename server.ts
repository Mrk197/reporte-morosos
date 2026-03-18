import express, { Request, Response } from 'express';
import { spawn } from 'child_process';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { platform } from 'os';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();
const PORT = 4000;

const RESULTS_DIR = path.join(__dirname, 'results');
const HISTORY_FILE = path.join(RESULTS_DIR, 'history.json');
const LATEST_RESULT_FILE = path.join(RESULTS_DIR, 'latest.json');

// Usar el binario local de tsx (compatible con todos los sistemas operativos)
const TSX_CMD = platform() === 'win32'
  ? path.join(__dirname, 'node_modules', '.bin', 'tsx.cmd')
  : path.join(__dirname, 'node_modules', '.bin', 'tsx');

// Crear directorio de resultados si no existe
if (!fs.existsSync(RESULTS_DIR)) {
  fs.mkdirSync(RESULTS_DIR, { recursive: true });
}

// Middleware
app.use(express.json());
app.use(express.static(__dirname));

// ============================================================
// API Endpoints
// ============================================================

// GET: Servir la página principal
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// GET: Obtener últimos resultados
app.get('/api/results', (req, res) => {
  try {
    if (fs.existsSync(LATEST_RESULT_FILE)) {
      const data = fs.readFileSync(LATEST_RESULT_FILE, 'utf-8');
      res.json(JSON.parse(data));
    } else {
      res.json(null);
    }
  } catch (err) {
    res.status(500).json({ error: 'Error leyendo resultados' });
  }
});

// GET: Obtener histórico de ejecuciones
app.get('/api/history', (req, res) => {
  try {
    if (fs.existsSync(HISTORY_FILE)) {
      const data = fs.readFileSync(HISTORY_FILE, 'utf-8');
      res.json(JSON.parse(data) || []);
    } else {
      res.json([]);
    }
  } catch (err) {
    res.status(500).json({ error: 'Error leyendo histórico' });
  }
});

// POST: Ejecutar el script
app.post('/api/run-script', (req, res) => {
  const { 
    format = 'json',
    pages = '20',
    perfil = '2026',
    categoria = 'retirar',
    diasInstalacion = '',
    desdeFecha = '',
    hastaFecha = ''
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

  // Ejecutar el script con tsx
  const args = [
    'new-clients-unpaid.ts',
    '--format', String(format),
    '--pages', String(pages),
    '--perfil', String(perfil),
    '--categoria', String(categoria),
  ];

  // Agregar parámetros opcionales de fecha
  if (diasInstalacion) {
    args.push('--dias-instalacion', String(diasInstalacion));
  }
  if (desdeFecha) {
    args.push('--desde-fecha', String(desdeFecha));
  }
  if (hastaFecha) {
    args.push('--hasta-fecha', String(hastaFecha));
  }

  const child = spawn(TSX_CMD, args, {
    cwd: __dirname,
    stdio: ['pipe', 'pipe', 'pipe'],
    shell: true,
  });

  let stdoutData = '';
  let stderrData = '';

  child.stdout?.on('data', (data) => {
    const text = data.toString();
    stdoutData += text;
    
    // Enviar líneas de progreso al cliente
    const lines = text.split('\n');
    for (const line of lines) {
      if (line.trim() && !line.startsWith('{')) {
        sendEvent({ status: 'procesando', linea: line });
      }
    }
  });

  child.stderr?.on('data', (data) => {
    stderrData += data.toString();
    sendEvent({ status: 'error', linea: data.toString() });
  });

  child.on('close', (code) => {
    if (code === 0) {
      try {
        // Parsear el JSON del stdout
        const jsonMatch = stdoutData.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
          const result = JSON.parse(jsonMatch[0]);

          // Guardar resultado actual
          fs.writeFileSync(LATEST_RESULT_FILE, JSON.stringify(result, null, 2));

          // Guardar en histórico
          const history = fs.existsSync(HISTORY_FILE)
            ? JSON.parse(fs.readFileSync(HISTORY_FILE, 'utf-8'))
            : [];
          
          // Determinar filtro de fecha usado
          let filtroFecha = 'Sin filtro';
          if (diasInstalacion) {
            filtroFecha = `Últimos ${diasInstalacion} días`;
          } else if (desdeFecha || hastaFecha) {
            const desde = desdeFecha || 'inicio';
            const hasta = hastaFecha || 'hoy';
            filtroFecha = `${desde} a ${hasta}`;
          }
          
          history.push({
            timestamp: new Date().toISOString(),
            formato: format,
            perfil: perfil,
            categoria: categoria,
            filtroFecha: filtroFecha,
            resumen: result.resumen,
            totales: {
              retirarModem: result.retirarModem?.length || 0,
              suspendidosSinPago: result.suspendidosSinPago?.length || 0,
              sinFacturaAun: result.sinFacturaAun?.length || 0,
            },
          });

          fs.writeFileSync(HISTORY_FILE, JSON.stringify(history, null, 2));

          sendEvent({
            status: 'completo',
            resultado: result,
            mensaje: 'Script ejecutado exitosamente',
          });
        } else {
          sendEvent({ status: 'error', mensaje: 'No se pudo parsear el resultado' });
        }
      } catch (err) {
        sendEvent({ status: 'error', mensaje: `Error: ${err instanceof Error ? err.message : 'Error desconocido'}` });
      }
    } else {
      sendEvent({ status: 'error', mensaje: `Script falló con código ${code}\n${stderrData}` });
      console.error('STDERR:', stderrData);
      console.error('STDOUT:', stdoutData);
    }
    res.end();
  });
});

// GET: Descargar Excel/CSV
app.get('/api/download/:format', (req, res) => {
  const { format } = req.params;
  const { pages = '20', perfil = '2026', categoria = 'retirar', diasInstalacion = '', desdeFecha = '', hastaFecha = '' } = req.query;

  const args = [
    'new-clients-unpaid.ts',
    '--format', String(format),
    '--pages', String(pages),
    '--perfil', String(perfil),
    '--categoria', String(categoria),
  ];

  // Agregar parámetros opcionales de fecha
  if (diasInstalacion) {
    args.push('--dias-instalacion', String(diasInstalacion));
  }
  if (desdeFecha) {
    args.push('--desde-fecha', String(desdeFecha));
  }
  if (hastaFecha) {
    args.push('--hasta-fecha', String(hastaFecha));
  }

  const child = spawn(TSX_CMD, args, {
    cwd: __dirname,
    stdio: ['pipe', 'pipe', 'pipe'],
    shell: true,
  });

  let stdoutData = '';
  let stderrData = '';

  child.stdout?.on('data', (data) => {
    stdoutData += data.toString();
  });

  child.stderr?.on('data', (data) => {
    stderrData += data.toString();
  });

  child.on('close', (code) => {
    if (code === 0) {
      // Buscar el archivo generado (con timestamp YYYYMMDDHHMMSS)
      let baseFilename = '';
      
      if (categoria === 'sinFactura') {
        baseFilename = 'sin-factura-aun';
      } else if (categoria === 'todos') {
        baseFilename = 'todos-sin-pagar';
      } else {
        baseFilename = 'retirar-modem';
      }

      const fileExtension = format === 'excel' ? 'xlsx' : format;
      const filepath = path.join(__dirname, `${baseFilename}.${fileExtension}`);

      if (fs.existsSync(filepath)) {
        res.download(filepath, `${baseFilename}.${fileExtension}`);
      } else {
        res.status(404).json({ error: `Archivo no encontrado: ${baseFilename}.${fileExtension}` });
      }
    } else {
      res.status(500).json({ error: `Error ejecutando script: ${stderrData}` });
    }
  });
});

// Iniciar servidor en 0.0.0.0 para permitir acceso remoto y ngrok
app.listen(PORT, '0.0.0.0', () => {
  console.log(`\n✅ Servidor iniciado en http://0.0.0.0:${PORT}`);
  console.log(`📊 Acceso local: http://localhost:${PORT}`);
  console.log(`🌐 Para acceso remoto, usa ngrok: npx ngrok http ${PORT}\n`);
});
