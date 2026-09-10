const test = require('node:test');
const assert = require('node:assert/strict');

const app = require('./server');

test('normalizes a small inline logo in the Cajas email config', () => {
  const config = app.normalizeCajasEmailConfig({
    recipients: ['destino@example.com'],
    cc: [],
    bcc: [],
    subject: 'Reporte',
    bodyHtml: '<p>Listo</p>',
    logoData: 'data:image/png;base64,bG9nbw==',
    logoName: 'logo.png',
    logoContentType: 'image/png',
  });

  assert.equal(config.logoData, 'data:image/png;base64,bG9nbw==');
  assert.equal(config.logoName, 'logo.png');
  assert.equal(config.logoContentType, 'image/png');
});

test('rejects unsupported inline logo formats', () => {
  assert.throws(
    () => app.normalizeCajasEmailConfig({
      recipients: ['destino@example.com'],
      cc: [],
      bcc: [],
      subject: 'Reporte',
      bodyHtml: '<p>Listo</p>',
      logoData: 'data:image/svg+xml;base64,PHN2Zy8+',
      logoName: 'logo.svg',
      logoContentType: 'image/svg+xml',
    }),
    /formato del logo|logo/i,
  );
});
