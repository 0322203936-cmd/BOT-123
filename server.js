require('dotenv').config();

const crypto = require('node:crypto');
const path = require('node:path');
const zlib = require('node:zlib');
const express = require('express');
const helmet = require('helmet');

const app = express();
const port = Number(process.env.PORT) || 3000;
const owner = process.env.GITHUB_OWNER || '0322203936-cmd';
const branch = process.env.GITHUB_BRANCH || 'main';
const githubToken = process.env.GITHUB_TOKEN || '';
const appPassword = process.env.APP_PASSWORD || '';

const workflows = {
  galleria: {
    owner,
    repo: process.env.GALLERIA_GITHUB_REPO || 'BOT-GALLERIA',
    branch,
    file: 'bot.yml',
    name: 'Reporte Galleria',
    description: 'Descarga el reporte de Galleria Farms y lo carga automáticamente en Posco.',
    schedule: 'Diario · 4:00 AM',
  },
  cancelaciones: {
    owner,
    repo: process.env.GALLERIA_GITHUB_REPO || 'BOT-GALLERIA',
    branch,
    file: 'cancelaciones.yml',
    name: 'Cancelaciones',
    description: 'Consulta solicitudes pendientes y actualiza el reporte acumulado de cancelaciones.',
    schedule: 'Lun–Sáb · 6:00 AM, 11:40 AM y 3:00 PM',
  },
  pegarData: {
    owner,
    repo: process.env.PEGAR_DATA_GITHUB_REPO || 'BOT-123',
    branch,
    file: 'pegar-data.yml',
    name: 'Pegar Data',
    description: 'Descarga datos desde Posco y actualiza el archivo de SharePoint.',
    schedule: 'Ejecución manual',
  },
  inventario: {
    owner,
    repo: process.env.PEGAR_DATA_GITHUB_REPO || 'BOT-123',
    branch,
    file: 'inventario.yml',
    name: 'Inventario',
    description: 'Descarga Inventario General de Posco y actualiza cuatro columnas en SharePoint.',
    schedule: 'Ejecución manual',
  },
  reunion: {
    owner,
    repo: process.env.PEGAR_DATA_GITHUB_REPO || 'BOT-123',
    branch,
    file: 'procesar-reunion.yml',
    name: 'Procesar Reunión',
    description: 'Avanza la fecha y recorre los campos Cor de la hoja Reunion.',
    schedule: 'Ejecución manual',
  },
  dataProy: {
    owner,
    repo: process.env.DATA_PROY_GITHUB_REPO || 'BOT-123',
    branch,
    file: 'data-proy.yml',
    name: 'Data Proy',
    description: 'Cruza proyecciones del Plan de Cosecha hacia Requerimientos.',
    schedule: 'Ejecución manual',
  },
  ainventario: {
    owner,
    repo: process.env.PEGAR_DATA_GITHUB_REPO || 'BOT-123',
    branch,
    file: 'ainventario.yml',
    name: 'AINVENTARIO',
    description: 'Sincroniza y suma inventario de Posco al requerimiento protegiendo la estructura.',
    schedule: 'Ejecución manual',
  },
  dataReq: {
    owner,
    repo: process.env.DATA_REQ_GITHUB_REPO || 'BOT-123',
    branch,
    file: 'data-req.yml',
    name: 'Data Req',
    description: 'Descarga reporte de Posco y lo pega en la hoja DataReq omitiendo columna N.',
    schedule: 'Ejecución manual',
  },
  cajas: {
    owner,
    repo: process.env.CAJAS_GITHUB_REPO || 'BOT-123',
    branch,
    file: 'inventory-boxes.yml',
    name: 'Subir XLS Cajas',
    description: 'Vacía el inventario de Kometsales y sube el XLS actualizado desde SharePoint.',
    schedule: 'Ejecución manual',
  },
};

const lastDispatch = new Map();
const cajasEmailVariable = 'CAJAS_EMAIL_CONFIG';
const graphUrl = 'https://graph.microsoft.com/v1.0';
const cajasSharePointUrl = process.env.CAJAS_SHAREPOINT_URL || 'https://pacificafarms.sharepoint.com/:x:/r/sites/requerimientovsproyeccion/_layouts/15/Doc.aspx?sourcedoc=%7B432E0F6F-229A-4635-A25A-A049DC537883%7D&file=Inventory%20Upload%20Boxes%2009092026.xlsx&action=default&mobileredirect=true';
const maxInlineLogoBytes = 24 * 1024;
const maxSharePointLogoBytes = 10 * 1024 * 1024;
const maxCajasPdfBytes = 10 * 1024 * 1024;
const inlineLogoTypes = new Set(['image/png', 'image/jpeg', 'image/gif', 'image/bmp']);
const logoExtensions = {
  'image/png': '.png',
  'image/jpeg': '.jpg',
  'image/gif': '.gif',
  'image/bmp': '.bmp',
};
const defaultCajasEmailConfig = {
  recipients: [],
  cc: [],
  bcc: [],
  subject: 'Reporte de inventario de cajas',
  bodyHtml: '<p>Hola,</p><p>Adjunto encontrarás el reporte actualizado de inventario de cajas.</p><p>Saludos.</p>',
  logoUrl: '',
  logoData: '',
  logoName: '',
  logoContentType: '',
  logoSharePoint: null,
};

const reportMatchers = {
  galleria: [
    (name) => /(^|\/)reporte_galleria_.*\.(?:xlsx?|xlsm)$/i.test(name),
  ],
  cancelaciones: [
    (name) => /(^|\/)reporte_cancelaciones_pendientes\.csv$/i.test(name),
  ],
  pegarData: [
    (name) => /(^|\/)[^/]*_formateado\.xlsx$/i.test(name),
    (name) => /(^|\/)[^/]*\.xlsx$/i.test(name),
    (name) => /(^|\/)[^/]*\.xls$/i.test(name),
    (name) => /_actualizado\.xlsm$/i.test(name),
  ],
  inventario: [
    (name) => /(^|\/)[^/]*\.xlsx$/i.test(name),
  ],
  dataProy: [
    (name) => /(^|\/)[^/]*\.(?:xlsx?|xlsm|csv)$/i.test(name),
  ],
  ainventario: [
    (name) => /(^|\/)[^/]*\.xlsx$/i.test(name),
  ],
  dataReq: [
    (name) => /(^|\/)[^/]*\.xlsx$/i.test(name),
  ],
  cajas: [
    (name) => /(^|\/)inventory-upload-boxes-.*\.xlsx$/i.test(name),
  ],
};

const artifactNameMatchers = {
  galleria: [(name) => name.startsWith('reporte-galleria-')],
  cancelaciones: [(name) => name.startsWith('reporte-cancelaciones-')],
  pegarData: [
    (name) => name.startsWith('reportes-pegar-data-'),
    (name) => name.startsWith('capturas-pegar-data-'),
  ],
  inventario: [
    (name) => name.startsWith('reportes-inventario-'),
    (name) => name.startsWith('capturas-inventario-'),
  ],
  dataProy: [(name) => name.startsWith('logs-data-proy-')],
  ainventario: [
    (name) => name.startsWith('reportes-ainventario-'),
    (name) => name === 'evidencias-posco',
  ],
  dataReq: [
    (name) => name.startsWith('reportes-data-req-'),
    (name) => name === 'evidencias-posco-datareq',
  ],
  cajas: [(name) => name.startsWith('reportes-inventory-boxes-')],
};

app.disable('x-powered-by');
app.use(helmet({ contentSecurityPolicy: false }));
app.use(
  '/api/workflows/cajas/email-logo',
  express.raw({ type: [...inlineLogoTypes], limit: `${maxSharePointLogoBytes}b` }),
);
app.use(
  '/api/workflows/cajas/email-pdf',
  express.raw({ type: 'application/pdf', limit: `${maxCajasPdfBytes}b` }),
);
app.use('/api/workflows/cajas/email-config', express.json({ limit: '100kb' }));
app.use(express.json({ limit: '10kb' }));

function passwordsMatch(received) {
  if (!appPassword) return true;
  const expected = Buffer.from(appPassword);
  const actual = Buffer.from(received || '');
  return expected.length === actual.length && crypto.timingSafeEqual(expected, actual);
}

function authenticate(req, res, next) {
  if (!passwordsMatch(req.get('X-App-Password'))) {
    return res.status(401).json({ message: 'La contraseña no es correcta.' });
  }
  next();
}

function requireWriteAuthentication(req, res, next) {
  if (!appPassword) {
    return res.status(503).json({ message: 'Falta configurar APP_PASSWORD en Render para guardar esta configuración.' });
  }
  return authenticate(req, res, next);
}

async function githubRequest(endpoint, options = {}) {
  if (!githubToken) {
    const error = new Error('Falta configurar GITHUB_TOKEN en Render.');
    error.status = 503;
    throw error;
  }

  const response = await fetch(`https://api.github.com${endpoint}`, {
    ...options,
    headers: {
      Accept: 'application/vnd.github+json',
      Authorization: `Bearer ${githubToken}`,
      'X-GitHub-Api-Version': '2022-11-28',
      'User-Agent': 'Atajos-Globales',
      ...(options.headers || {}),
    },
  });

  if (!response.ok) {
    let detail = '';
    try {
      const body = await response.json();
      detail = body.message ? ` ${body.message}` : '';
    } catch {}
    const error = new Error(`GitHub rechazó la solicitud (${response.status}).${detail}`);
    error.status = response.status === 401 || response.status === 403 ? 502 : response.status;
    throw error;
  }

  if (response.status === 204) return null;
  return response.json();
}

async function githubBinary(endpoint) {
  if (!githubToken) {
    const error = new Error('Falta configurar GITHUB_TOKEN en Render.');
    error.status = 503;
    throw error;
  }

  const response = await fetch(`https://api.github.com${endpoint}`, {
    headers: {
      Accept: 'application/vnd.github+json',
      Authorization: `Bearer ${githubToken}`,
      'X-GitHub-Api-Version': '2022-11-28',
      'User-Agent': 'Atajos-Globales',
    },
  });

  if (!response.ok) {
    const error = new Error(`GitHub rechazó la descarga (${response.status}).`);
    error.status = response.status === 401 || response.status === 403 ? 502 : response.status;
    throw error;
  }

  return Buffer.from(await response.arrayBuffer());
}

function zipEntries(archive) {
  const endOfCentralDirectory = Buffer.from([0x50, 0x4b, 0x05, 0x06]);
  const minimumOffset = Math.max(0, archive.length - 65_557);
  let endOffset = -1;
  for (let offset = archive.length - 22; offset >= minimumOffset; offset -= 1) {
    if (archive.subarray(offset, offset + 4).equals(endOfCentralDirectory)) {
      endOffset = offset;
      break;
    }
  }
  if (endOffset < 0) throw new Error('El artifact descargado no es un ZIP válido.');

  const entryCount = archive.readUInt16LE(endOffset + 10);
  const centralDirectoryOffset = archive.readUInt32LE(endOffset + 16);
  const entries = [];
  let offset = centralDirectoryOffset;

  for (let index = 0; index < entryCount; index += 1) {
    if (archive.readUInt32LE(offset) !== 0x02014b50) {
      throw new Error('El índice del artifact ZIP está dañado.');
    }
    const compression = archive.readUInt16LE(offset + 10);
    const compressedSize = archive.readUInt32LE(offset + 20);
    const uncompressedSize = archive.readUInt32LE(offset + 24);
    const nameLength = archive.readUInt16LE(offset + 28);
    const extraLength = archive.readUInt16LE(offset + 30);
    const commentLength = archive.readUInt16LE(offset + 32);
    const localHeaderOffset = archive.readUInt32LE(offset + 42);
    const name = archive.toString('utf8', offset + 46, offset + 46 + nameLength);
    entries.push({
      name,
      compression,
      compressedSize,
      uncompressedSize,
      localHeaderOffset,
      directory: name.endsWith('/'),
    });
    offset += 46 + nameLength + extraLength + commentLength;
  }

  return entries;
}

function extractZipEntry(archive, entry) {
  if (entry.directory) return null;
  const localOffset = entry.localHeaderOffset;
  if (archive.readUInt32LE(localOffset) !== 0x04034b50) {
    throw new Error('La entrada del artifact ZIP está dañada.');
  }
  const nameLength = archive.readUInt16LE(localOffset + 26);
  const extraLength = archive.readUInt16LE(localOffset + 28);
  const dataStart = localOffset + 30 + nameLength + extraLength;
  const compressed = archive.subarray(dataStart, dataStart + entry.compressedSize);
  let data;
  if (entry.compression === 0) data = compressed;
  else if (entry.compression === 8) data = zlib.inflateRawSync(compressed);
  else throw new Error(`Compresión ZIP no soportada para ${entry.name}.`);
  if (data.length !== entry.uncompressedSize) {
    throw new Error(`El archivo ${entry.name} quedó incompleto al extraerlo.`);
  }
  return data;
}

function sharePointId(value) {
  return `u!${Buffer.from(value, 'utf8').toString('base64url')}`;
}

function sharePointError(response, action) {
  return response.text().then((detail) => {
    const error = new Error(`SharePoint no pudo ${action} (HTTP ${response.status}). ${detail.slice(0, 500)}`);
    error.status = response.status >= 400 && response.status < 500 ? 502 : 503;
    throw error;
  });
}

function requiredSharePointEnv(name) {
  const value = process.env[name]?.trim();
  if (!value) {
    const error = new Error(`Falta configurar ${name} en Render para cargar el logo de Cajas.`);
    error.status = 503;
    throw error;
  }
  return value;
}

async function sharePointToken() {
  const tenantId = requiredSharePointEnv('SHAREPOINT_TENANT_ID');
  const response = await fetch(
    `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: requiredSharePointEnv('SHAREPOINT_CLIENT_ID'),
        client_secret: requiredSharePointEnv('SHAREPOINT_CLIENT_SECRET'),
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials',
      }),
    },
  );
  if (!response.ok) return sharePointError(response, 'autenticar la carga del logo');
  const payload = await response.json();
  if (!payload.access_token) {
    const error = new Error('SharePoint no devolvió un token para cargar el logo de Cajas.');
    error.status = 503;
    throw error;
  }
  return payload.access_token;
}

async function uploadCajasLogo(buffer, contentType) {
  if (!Buffer.isBuffer(buffer) || !buffer.length) {
    const error = new Error('Selecciona un archivo de logo válido.');
    error.status = 400;
    throw error;
  }
  if (buffer.length > maxSharePointLogoBytes) {
    const error = new Error('El logo debe pesar como máximo 10 MB.');
    error.status = 400;
    throw error;
  }
  if (!inlineLogoTypes.has(contentType)) {
    const error = new Error('El logo debe ser PNG, JPG, GIF o BMP.');
    error.status = 415;
    throw error;
  }

  const token = await sharePointToken();
  const itemResponse = await fetch(`${graphUrl}/shares/${sharePointId(cajasSharePointUrl)}/driveItem`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!itemResponse.ok) return sharePointError(itemResponse, 'localizar el archivo de Cajas');
  const item = await itemResponse.json();
  const driveId = item.parentReference?.driveId;
  const parentId = item.parentReference?.id;
  if (!driveId || !parentId) {
    const error = new Error('SharePoint no devolvió la carpeta del archivo de Cajas.');
    error.status = 502;
    throw error;
  }

  const filename = `cajas-email-logo-${crypto.randomUUID()}${logoExtensions[contentType]}`;
  const uploadResponse = await fetch(
    `${graphUrl}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentId)}:/${encodeURIComponent(filename)}:/content`,
    {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': contentType,
        'Content-Length': String(buffer.length),
      },
      body: buffer,
    },
  );
  if (!uploadResponse.ok) return sharePointError(uploadResponse, 'guardar el logo');
  const uploaded = await uploadResponse.json();
  if (!uploaded.id) {
    const error = new Error('SharePoint guardó el logo sin devolver su identificador.');
    error.status = 502;
    throw error;
  }
  return {
    driveId,
    itemId: uploaded.id,
    name: filename,
    contentType,
  };
}

async function uploadCajasPdf(buffer, contentType) {
  if (!Buffer.isBuffer(buffer) || !buffer.length) {
    const error = new Error('Selecciona un PDF válido.'); error.status = 400; throw error;
  }
  if (contentType !== 'application/pdf' || buffer.subarray(0, 5).toString() !== '%PDF-') {
    const error = new Error('El archivo debe ser PDF.'); error.status = 415; throw error;
  }
  if (buffer.length > maxCajasPdfBytes) {
    const error = new Error('El PDF debe pesar como máximo 10 MB.'); error.status = 400; throw error;
  }
  const token = await sharePointToken();
  const itemResponse = await fetch(`${graphUrl}/shares/${sharePointId(cajasSharePointUrl)}/driveItem`, { headers: { Authorization: `Bearer ${token}` } });
  if (!itemResponse.ok) return sharePointError(itemResponse, 'localizar el archivo de Cajas');
  const item = await itemResponse.json();
  const driveId = item.parentReference?.driveId;
  const parentId = item.parentReference?.id;
  if (!driveId || !parentId) { const error = new Error('SharePoint no devolvió la carpeta del archivo de Cajas.'); error.status = 502; throw error; }
  const filename = `cajas-email-attachment-${crypto.randomUUID()}.pdf`;
  const uploadResponse = await fetch(`${graphUrl}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentId)}:/${encodeURIComponent(filename)}:/content`, {
    method: 'PUT', headers: { Authorization: `Bearer ${token}`, 'Content-Type': contentType, 'Content-Length': String(buffer.length) }, body: buffer,
  });
  if (!uploadResponse.ok) return sharePointError(uploadResponse, 'guardar el PDF');
  const uploaded = await uploadResponse.json();
  if (!uploaded.id) { const error = new Error('SharePoint guardó el PDF sin devolver su identificador.'); error.status = 502; throw error; }
  return { driveId, itemId: uploaded.id, name: filename, contentType };
}

async function deleteCajasPdf(value) {
  if (!value || typeof value !== 'object' || value.contentType !== 'application/pdf' || !/^cajas-email-attachment-[0-9a-f-]{36}\.pdf$/i.test(value.name || '')) {
    const error = new Error('La referencia del PDF temporal no es válida.'); error.status = 400; throw error;
  }
  const token = await sharePointToken();
  const response = await fetch(`${graphUrl}/drives/${encodeURIComponent(value.driveId)}/items/${encodeURIComponent(value.itemId)}`, { method: 'DELETE', headers: { Authorization: `Bearer ${token}` } });
  if (!response.ok && response.status !== 404) return sharePointError(response, 'eliminar el PDF temporal');
}

async function deleteCajasLogo(value) {
  const logo = normalizeCajasLogoSharePoint(value);
  if (!logo || !/^cajas-email-logo-[0-9a-f-]{36}\.(png|jpg|gif|bmp)$/i.test(logo.name)) {
    const error = new Error('La referencia del logo temporal no es válida.');
    error.status = 400;
    throw error;
  }

  const token = await sharePointToken();
  const response = await fetch(
    `${graphUrl}/drives/${encodeURIComponent(logo.driveId)}/items/${encodeURIComponent(logo.itemId)}`,
    {
      method: 'DELETE',
      headers: { Authorization: `Bearer ${token}` },
    },
  );
  if (!response.ok && response.status !== 404) return sharePointError(response, 'eliminar el logo temporal');
}

function emailList(value, field) {
  if (!Array.isArray(value)) {
    const error = new Error(`El campo ${field} debe ser una lista de correos.`);
    error.status = 400;
    throw error;
  }
  if (value.length > 100) {
    const error = new Error(`El campo ${field} no puede tener más de 100 correos.`);
    error.status = 400;
    throw error;
  }
  const result = [];
  for (const item of value) {
    if (typeof item !== 'string') {
      const error = new Error(`El campo ${field} contiene un correo inválido.`);
      error.status = 400;
      throw error;
    }
    const address = item.trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(address)) {
      const error = new Error(`El campo ${field} contiene un correo inválido.`);
      error.status = 400;
      throw error;
    }
    if (!result.some((existing) => existing.toLowerCase() === address.toLowerCase())) {
      result.push(address);
    }
  }
  return result;
}

function normalizeCajasLogoSharePoint(value) {
  if (value === undefined || value === null || value === '') return null;
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    const error = new Error('La referencia del logo en SharePoint no es válida.');
    error.status = 400;
    throw error;
  }
  const driveId = typeof value.driveId === 'string' ? value.driveId.trim() : '';
  const itemId = typeof value.itemId === 'string' ? value.itemId.trim() : '';
  const name = typeof value.name === 'string' ? value.name.trim() : '';
  const contentType = typeof value.contentType === 'string' ? value.contentType.trim().toLowerCase() : '';
  if (!driveId || !itemId || !name || !inlineLogoTypes.has(contentType)) {
    const error = new Error('La referencia del logo en SharePoint está incompleta.');
    error.status = 400;
    throw error;
  }
  if (name.length > 100 || /[\\/\r\n]/.test(name)) {
    const error = new Error('El nombre del logo no es válido.');
    error.status = 400;
    throw error;
  }
  return { driveId, itemId, name, contentType };
}

function normalizeCajasPdfSharePoint(value) {
  if (value === undefined || value === null || value === '') return null;
  if (!value || typeof value !== 'object' || Array.isArray(value)) { const error = new Error('La referencia del PDF en SharePoint no es válida.'); error.status = 400; throw error; }
  const driveId = typeof value.driveId === 'string' ? value.driveId.trim() : '';
  const itemId = typeof value.itemId === 'string' ? value.itemId.trim() : '';
  const name = typeof value.name === 'string' ? value.name.trim() : '';
  if (!driveId || !itemId || !/^cajas-email-attachment-[0-9a-f-]{36}\.pdf$/i.test(name) || value.contentType !== 'application/pdf') { const error = new Error('La referencia del PDF en SharePoint está incompleta.'); error.status = 400; throw error; }
  return { driveId, itemId, name, contentType: 'application/pdf' };
}

function normalizeCajasEmailConfig(value) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    const error = new Error('La configuración del correo debe ser un objeto.');
    error.status = 400;
    throw error;
  }
  const recipients = emailList(value.recipients, 'recipients');
  if (!recipients.length) {
    const error = new Error('Agrega al menos un destinatario.');
    error.status = 400;
    throw error;
  }
  const cc = emailList(value.cc || [], 'cc');
  const bcc = emailList(value.bcc || [], 'bcc');
  const subject = typeof value.subject === 'string' ? value.subject.trim() : '';
  const bodyHtml = typeof value.bodyHtml === 'string' ? value.bodyHtml : '';
  const logoUrl = typeof value.logoUrl === 'string' ? value.logoUrl.trim() : '';
  const logoData = typeof value.logoData === 'string' ? value.logoData.trim() : '';
  let logoName = typeof value.logoName === 'string' ? value.logoName.trim() : '';
  const logoSharePoint = normalizeCajasLogoSharePoint(value.logoSharePoint);
  const pdfSharePoint = normalizeCajasPdfSharePoint(value.pdfSharePoint);
  if (!subject || subject.length > 200) {
    const error = new Error('El asunto es obligatorio y debe tener hasta 200 caracteres.');
    error.status = 400;
    throw error;
  }
  if (!bodyHtml.trim() || bodyHtml.length > 100_000) {
    const error = new Error('El contenido del correo es obligatorio y demasiado grande.');
    error.status = 400;
    throw error;
  }
  if (/<\/?script\b|\son\w+\s*=|javascript:/i.test(bodyHtml)) {
    const error = new Error('El contenido del correo incluye HTML no permitido.');
    error.status = 400;
    throw error;
  }
  if (logoUrl && !/^https?:\/\//i.test(logoUrl)) {
    const error = new Error('La URL del logo debe comenzar con http:// o https://.');
    error.status = 400;
    throw error;
  }
  let logoContentType = '';
  if (logoData && !logoSharePoint) {
    const match = logoData.match(/^data:(image\/(?:png|jpeg|gif|bmp));base64,([A-Za-z0-9+/]*={0,2})$/i);
    if (!match || match[2].length % 4 !== 0) {
      const error = new Error('El logo debe ser una imagen PNG, JPG, GIF o BMP.');
      error.status = 400;
      throw error;
    }
    logoContentType = match[1].toLowerCase();
    if (!inlineLogoTypes.has(logoContentType)) {
      const error = new Error('El formato del logo no está permitido.');
      error.status = 400;
      throw error;
    }
    const logoBytes = Buffer.from(match[2], 'base64');
    if (!logoBytes.length || logoBytes.length > maxInlineLogoBytes) {
      const error = new Error('El logo debe pesar como máximo 24 KB.');
      error.status = 400;
      throw error;
    }
    if (!logoName) {
      logoName = {
        'image/png': 'logo.png',
        'image/jpeg': 'logo.jpg',
        'image/gif': 'logo.gif',
        'image/bmp': 'logo.bmp',
      }[logoContentType];
    }
    if (logoName.length > 100 || /[\\/\r\n]/.test(logoName)) {
      const error = new Error('El nombre del logo no es válido.');
      error.status = 400;
      throw error;
    }
  } else {
    logoName = '';
  }
  const config = {
    recipients,
    cc,
    bcc,
    subject,
    bodyHtml,
    logoUrl: logoData || logoSharePoint ? '' : logoUrl,
    logoData: logoSharePoint ? '' : logoData,
    logoName: logoSharePoint ? '' : logoName,
    logoContentType: logoSharePoint ? '' : logoContentType,
    logoSharePoint,
    pdfSharePoint,
  };
  if (Buffer.byteLength(JSON.stringify(config), 'utf8') > 40_000) {
    const error = new Error('La configuración del correo es demasiado grande para GitHub Actions.');
    error.status = 400;
    throw error;
  }
  return config;
}

async function readCajasEmailConfig(workflow) {
  try {
    const variable = await githubRequest(
      `/repos/${encodeURIComponent(workflow.owner)}/${encodeURIComponent(workflow.repo)}/actions/variables/${encodeURIComponent(cajasEmailVariable)}`,
    );
    const config = normalizeCajasEmailConfig(JSON.parse(variable.value));
    return { config, configured: true };
  } catch (error) {
    if (error.status === 404) {
      return { config: { ...defaultCajasEmailConfig }, configured: false };
    }
    if (error instanceof SyntaxError) {
      error.message = 'La variable CAJAS_EMAIL_CONFIG no contiene un JSON válido.';
      error.status = 502;
    }
    throw error;
  }
}

function findReportEntry(entries, matchers) {
  for (const matcher of matchers) {
    const file = entries.find((entry) => !entry.directory && matcher(entry.name));
    if (file) return file;
  }
  return null;
}

function reportContentType(filename) {
  const extension = path.extname(filename).toLowerCase();
  return {
    '.csv': 'text/csv; charset=utf-8',
    '.xls': 'application/vnd.ms-excel',
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.xlsm': 'application/vnd.ms-excel.sheet.macroEnabled.12',
  }[extension] || 'application/octet-stream';
}

async function latestReportFile(key, workflow) {
  const matchers = reportMatchers[key] || [];
  const artifactMatchers = artifactNameMatchers[key] || [];
  if (!matchers.length || !artifactMatchers.length) {
    const error = new Error('Este bot todavía no genera un archivo descargable.');
    error.status = 404;
    throw error;
  }

  const runsData = await githubRequest(
    `/repos/${encodeURIComponent(workflow.owner)}/${encodeURIComponent(workflow.repo)}/actions/workflows/${encodeURIComponent(workflow.file)}/runs?per_page=10`,
  );
  const runs = runsData.workflow_runs || [];
  const latest = runs[0];
  if (latest && (latest.status === 'queued' || latest.status === 'in_progress')) {
    const error = new Error('El bot todavía está ejecutándose; el reporte estará disponible al terminar.');
    error.status = 409;
    throw error;
  }

  for (const run of runs.filter((item) => item.status === 'completed')) {
    const artifactsData = await githubRequest(
      `/repos/${encodeURIComponent(workflow.owner)}/${encodeURIComponent(workflow.repo)}/actions/runs/${run.id}/artifacts?per_page=100`,
    );
    for (const artifactMatcher of artifactMatchers) {
      const artifact = (artifactsData.artifacts || []).find(
        (item) => !item.expired && artifactMatcher(item.name),
      );
      if (!artifact) continue;

      const archive = await githubBinary(
        `/repos/${encodeURIComponent(workflow.owner)}/${encodeURIComponent(workflow.repo)}/actions/artifacts/${artifact.id}/zip`,
      );
      const file = findReportEntry(zipEntries(archive), matchers);
      if (file) {
        return {
          data: extractZipEntry(archive, file),
          filename: path.basename(file.name),
          run,
        };
      }
    }
  }

  const error = new Error('No se encontró un reporte en las últimas ejecuciones de este bot.');
  error.status = 404;
  throw error;
}

function serializeRun(run) {
  if (!run) return null;
  return {
    id: run.id,
    status: run.status || 'unknown',
    conclusion: run.conclusion,
    createdAt: run.created_at,
    updatedAt: run.updated_at,
    url: run.html_url,
    event: run.event,
  };
}

async function latestRun(workflow) {
  const data = await githubRequest(
    `/repos/${encodeURIComponent(workflow.owner)}/${encodeURIComponent(workflow.repo)}/actions/workflows/${encodeURIComponent(workflow.file)}/runs?per_page=1`,
  );
  return serializeRun(data.workflow_runs?.[0]);
}

app.get('/api/health', (_req, res) => res.json({ ok: true }));

app.get('/api/config', (_req, res) => {
  res.json({ authRequired: Boolean(appPassword), configured: Boolean(githubToken) });
});

app.get('/api/workflows/cajas/email-config', authenticate, async (_req, res, next) => {
  try {
    const config = await readCajasEmailConfig(workflows.cajas);
    res.json(config);
  } catch (error) {
    next(error);
  }
});

app.post('/api/workflows/cajas/email-logo', requireWriteAuthentication, async (req, res, next) => {
  try {
    const contentType = (req.get('Content-Type') || '').split(';', 1)[0].trim().toLowerCase();
    if (!inlineLogoTypes.has(contentType)) {
      return res.status(415).json({ message: 'El logo debe ser PNG, JPG, GIF o BMP.' });
    }
    const logoSharePoint = await uploadCajasLogo(req.body, contentType);
    return res.json({ logoSharePoint });
  } catch (error) {
    return next(error);
  }
});

app.delete('/api/workflows/cajas/email-logo', requireWriteAuthentication, async (req, res, next) => {
  try {
    await deleteCajasLogo(req.body);
    return res.status(204).send();
  } catch (error) {
    return next(error);
  }
});

app.post('/api/workflows/cajas/email-pdf', requireWriteAuthentication, async (req, res, next) => {
  try {
    const contentType = (req.get('Content-Type') || '').split(';', 1)[0].trim().toLowerCase();
    const pdfSharePoint = await uploadCajasPdf(req.body, contentType);
    return res.json({ pdfSharePoint });
  } catch (error) { return next(error); }
});

app.delete('/api/workflows/cajas/email-pdf', requireWriteAuthentication, async (req, res, next) => {
  try { await deleteCajasPdf(req.body); return res.status(204).send(); }
  catch (error) { return next(error); }
});

app.put('/api/workflows/cajas/email-config', requireWriteAuthentication, async (req, res, next) => {
  try {
    const config = normalizeCajasEmailConfig(req.body);
    const variablePath = `/repos/${encodeURIComponent(workflows.cajas.owner)}/${encodeURIComponent(workflows.cajas.repo)}/actions/variables`;
    const variableBody = JSON.stringify({ name: cajasEmailVariable, value: JSON.stringify(config) });
    try {
      await githubRequest(`${variablePath}/${encodeURIComponent(cajasEmailVariable)}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: variableBody,
      });
    } catch (error) {
      if (error.status !== 404) throw error;
      await githubRequest(variablePath, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: variableBody,
      });
    }
    res.json({ config, message: 'Configuración del correo de Cajas guardada.' });
  } catch (error) {
    next(error);
  }
});

app.get('/api/workflows', authenticate, async (_req, res, next) => {
  try {
    const entries = await Promise.all(
      Object.entries(workflows).map(async ([key, workflow]) => {
        let run = await latestRun(workflow);
        const dispatchedAt = lastDispatch.get(key);
        const runCreatedAt = run ? new Date(run.createdAt).getTime() : 0;

        // GitHub puede tardar algunos segundos en publicar la nueva ejecución.
        // Conservamos un estado en cola para que la interfaz no vuelva al estado anterior.
        if (
          dispatchedAt &&
          Date.now() - dispatchedAt < 120_000 &&
          runCreatedAt < dispatchedAt - 2_000
        ) {
          const timestamp = new Date(dispatchedAt).toISOString();
          run = {
            id: 0,
            status: 'queued',
            conclusion: null,
            createdAt: timestamp,
            updatedAt: timestamp,
            url: '',
            event: 'workflow_dispatch',
          };
        }

        return {
          key,
          name: workflow.name,
          description: workflow.description,
          schedule: workflow.schedule,
          run,
        };
      }),
    );
    res.json({ workflows: entries });
  } catch (error) {
    next(error);
  }
});

app.post('/api/workflows/:key/dispatch', authenticate, async (req, res, next) => {
  try {
    const workflow = workflows[req.params.key];
    if (!workflow) return res.status(404).json({ message: 'Automatización no encontrada.' });

    const previous = lastDispatch.get(req.params.key) || 0;
    if (Date.now() - previous < 15_000) {
      return res.status(429).json({ message: 'Espera unos segundos antes de volver a ejecutar este bot.' });
    }

    const currentRun = await latestRun(workflow);
    if (currentRun && (currentRun.status === 'queued' || currentRun.status === 'in_progress')) {
      return res.status(409).json({ message: `${workflow.name} ya tiene una ejecución activa.` });
    }

    await githubRequest(
      `/repos/${encodeURIComponent(workflow.owner)}/${encodeURIComponent(workflow.repo)}/actions/workflows/${encodeURIComponent(workflow.file)}/dispatches`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ref: workflow.branch }),
      },
    );

    const dispatchedAt = Date.now();
    lastDispatch.set(req.params.key, dispatchedAt);
    res.status(202).json({
      message: `${workflow.name} fue enviado correctamente a GitHub.`,
      dispatchedAt: new Date(dispatchedAt).toISOString(),
    });
  } catch (error) {
    next(error);
  }
});

app.get('/api/workflows/:key/report', authenticate, async (req, res, next) => {
  try {
    const workflow = workflows[req.params.key];
    if (!workflow) return res.status(404).json({ message: 'Automatización no encontrada.' });

    const report = await latestReportFile(req.params.key, workflow);
    const safeFilename = report.filename.replace(/[\r\n"\\]/g, '_');
    res.setHeader('Content-Type', reportContentType(safeFilename));
    res.setHeader('Content-Length', report.data.length);
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="${safeFilename}"; filename*=UTF-8''${encodeURIComponent(safeFilename)}`,
    );
    return res.send(report.data);
  } catch (error) {
    return next(error);
  }
});

app.use('/api', (_req, res) => res.status(404).json({ message: 'Ruta no encontrada.' }));

const browserPath = path.join(__dirname, 'dist', 'atajos-globales', 'browser');
app.use(express.static(browserPath, { maxAge: '1d', index: false }));
app.use((req, res, next) => {
  if (req.method !== 'GET' || !req.accepts('html')) return next();
  res.sendFile(path.join(browserPath, 'index.html'));
});

app.use((error, _req, res, _next) => {
  console.error(error.message);
  res.status(error.status || 500).json({ message: error.message || 'Error interno del servidor.' });
});

if (require.main === module) {
  app.listen(port, '0.0.0.0', () => {
    console.log(`Atajos Globales disponible en el puerto ${port}`);
  });
}

app.normalizeCajasEmailConfig = normalizeCajasEmailConfig;
module.exports = app;
