require('dotenv').config();

const crypto = require('node:crypto');
const path = require('node:path');
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
    schedule: 'Diario · 7:00 AM, 12:00 PM y 4:00 PM',
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
};

const lastDispatch = new Map();

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
