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
};

const lastDispatch = new Map();

const reportMatchers = {
  galleria: [
    (name) => /(^|\/)reporte_galleria_.*\.(?:xlsx?|xlsm)$/i.test(name),
  ],
  cancelaciones: [
    (name) => /(^|\/)reporte_cancelaciones_pendientes\.csv$/i.test(name),
  ],
  pegarData: [
    (name) => /^artifacts\/reportes\/.*\.(?:xlsx?|xlsm)$/i.test(name),
    (name) => /_actualizado\.xlsm$/i.test(name),
  ],
  inventario: [
    (name) => /^artifacts\/inventario\/reportes\/.*\.xlsx$/i.test(name),
  ],
  dataProy: [
    (name) => /^artifacts\/sharepoint\/.*\.(?:xlsx?|xlsm|csv)$/i.test(name),
  ],
  ainventario: [
    (name) => /^artifacts\/inventario\/reportes\/.*\.xlsx$/i.test(name),
  ],
  dataReq: [
    (name) => /^artifacts\/reportes\/.*\.xlsx$/i.test(name),
  ],
};

const artifactNameMatchers = {
  galleria: (name) => name.startsWith('reporte-galleria-'),
  cancelaciones: (name) => name.startsWith('reporte-cancelaciones-'),
  pegarData: (name) => name.startsWith('capturas-pegar-data-'),
  inventario: (name) => name.startsWith('capturas-inventario-'),
  dataProy: (name) => name.startsWith('logs-data-proy-'),
  ainventario: (name) => name === 'evidencias-posco',
  dataReq: (name) => name === 'evidencias-posco-datareq',
};

app.disable('x-powered-by');
app.use(helmet({ contentSecurityPolicy: false }));
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

function unzipEntries(archive) {
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

  return entries.map((entry) => {
    if (entry.directory) return { ...entry, data: null };
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
    return { ...entry, data };
  });
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
  const artifactMatcher = artifactNameMatchers[key];
  if (!matchers.length || !artifactMatcher) {
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
    const artifact = (artifactsData.artifacts || []).find(
      (item) => !item.expired && artifactMatcher(item.name),
    );
    if (!artifact) continue;

    const archive = await githubBinary(
      `/repos/${encodeURIComponent(workflow.owner)}/${encodeURIComponent(workflow.repo)}/actions/artifacts/${artifact.id}/zip`,
    );
    const file = unzipEntries(archive).find(
      (entry) => !entry.directory && entry.data && matchers.some((matcher) => matcher(entry.name)),
    );
    if (file) {
      return {
        data: file.data,
        filename: path.basename(file.name),
        run,
      };
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

app.listen(port, '0.0.0.0', () => {
  console.log(`Atajos Globales disponible en el puerto ${port}`);
});

module.exports = app;
