/**
 * Cliente para MikroWisp API
 * Maneja conexiones HTTPS con certificados auto-firmados (IPs)
 */

import https from 'https';
import http from 'http';
import dotenv from 'dotenv';

dotenv.config();

const MIKROWISP_API_URL = process.env.MIKROWISP_API_URL || '';
const MIKROWISP_TOKEN = process.env.MIKROWISP_TOKEN || '';

function httpsPost(url: string, body: object): Promise<{ status: number; data: Record<string, unknown> }> {
  return new Promise((resolve, reject) => {
    const parsedUrl = new URL(url);
    const isHttps = parsedUrl.protocol === 'https:';
    const postData = JSON.stringify(body);

    const options = {
      hostname: parsedUrl.hostname,
      port: parsedUrl.port || (isHttps ? 443 : 80),
      path: parsedUrl.pathname + parsedUrl.search,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(postData),
      },
      ...(isHttps && { rejectUnauthorized: false }),
    };

    const lib = isHttps ? https : http;

    const req = lib.request(options, (res) => {
      let responseBody = '';
      res.on('data', (chunk) => { responseBody += chunk; });
      res.on('end', () => {
        try {
          const data = JSON.parse(responseBody);
          resolve({ status: res.statusCode || 200, data });
        } catch {
          reject(new Error(`Respuesta no JSON de MikroWisp: ${responseBody.substring(0, 200)}`));
        }
      });
    });

    req.on('error', (err) => reject(err));
    req.setTimeout(30000, () => {
      req.destroy();
      reject(new Error('Timeout: MikroWisp no respondió en 30 segundos'));
    });

    req.write(postData);
    req.end();
  });
}

export async function mikroWispPostRaw(endpoint: string, data: Record<string, unknown>): Promise<{ status: number; data: Record<string, unknown> }> {
  if (!MIKROWISP_API_URL || !MIKROWISP_TOKEN) {
    throw new Error('MikroWisp API no configurada. Crea un archivo .env con MIKROWISP_API_URL y MIKROWISP_TOKEN');
  }

  const url = `${MIKROWISP_API_URL}/api/v1/${endpoint}`;
  return httpsPost(url, { token: MIKROWISP_TOKEN, ...data });
}

export function extractFacturas(data: Record<string, unknown>): Array<Record<string, unknown>> {
  if (Array.isArray(data)) return data;
  if (data.facturas && Array.isArray(data.facturas)) return data.facturas;
  if (data.datos && Array.isArray(data.datos)) return data.datos;
  return [];
}
