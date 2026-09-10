import { TestBed } from '@angular/core/testing';
import { provideHttpClient } from '@angular/common/http';
import { HttpTestingController, provideHttpClientTesting } from '@angular/common/http/testing';
import { App } from './app';

describe('App', () => {
  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [App],
      providers: [provideHttpClient(), provideHttpClientTesting()],
    }).compileComponents();
  });

  it('should create the app', () => {
    const fixture = TestBed.createComponent(App);
    const app = fixture.componentInstance;
    expect(app).toBeTruthy();
  });

  it('should render the control panel', async () => {
    const fixture = TestBed.createComponent(App);
    const http = TestBed.inject(HttpTestingController);
    fixture.detectChanges();

    http.expectOne('/api/config').flush({ authRequired: false, configured: true });
    await Promise.resolve();
    http.expectOne('/api/workflows').flush({
      workflows: [
        {
          key: 'galleria',
          name: 'Reporte Galleria',
          description: 'Prueba',
          schedule: 'Diario',
          run: null,
        },
      ],
    });
    await new Promise((resolve) => setTimeout(resolve, 0));
    fixture.detectChanges();

    const compiled = fixture.nativeElement as HTMLElement;
    expect(compiled.querySelector('h1')?.textContent).toContain('Atajos Centro Floricultor');
    expect(compiled.querySelector('.run-button')?.textContent).toContain('Ejecutar');
    http.verify();
  });

  it('should show the email editor only for the Cajas workflow', async () => {
    const fixture = TestBed.createComponent(App);
    const http = TestBed.inject(HttpTestingController);
    fixture.detectChanges();

    http.expectOne('/api/config').flush({ authRequired: false, configured: true });
    await Promise.resolve();
    http.expectOne('/api/workflows').flush({
      workflows: [
        {
          key: 'galleria',
          name: 'Reporte Galleria',
          description: 'Prueba',
          schedule: 'Diario',
          run: null,
        },
        {
          key: 'cajas',
          name: 'Subir XLS Cajas',
          description: 'Prueba',
          schedule: 'Manual',
          run: null,
        },
      ],
    });
    await new Promise((resolve) => setTimeout(resolve, 0));
    fixture.detectChanges();

    const compiled = fixture.nativeElement as HTMLElement;
    expect(compiled.querySelectorAll('.email-button')).toHaveLength(1);
    expect(compiled.querySelector('.workflow-card:not(.senary) .email-button')).toBeNull();

    (compiled.querySelector('.email-button') as HTMLButtonElement).click();
    await new Promise((resolve) => setTimeout(resolve, 0));
    const configRequest = http.expectOne('/api/workflows/cajas/email-config');
    configRequest.flush({
      configured: false,
      config: {
        recipients: [],
        cc: [],
        bcc: [],
        subject: 'Reporte de inventario de cajas',
        bodyHtml: '<p>Listo</p>',
        logoData: '',
        logoName: '',
        logoContentType: '',
      },
    });
    await new Promise((resolve) => setTimeout(resolve, 0));
    fixture.detectChanges();

    expect(compiled.querySelector('input[type="file"]')).toBeTruthy();
    expect(compiled.querySelector('input[type="url"]')).toBeNull();
    http.verify();
  });

  it('should convert a selected logo into inline data', async () => {
    const fixture = TestBed.createComponent(App);
    const app = fixture.componentInstance as any;
    const input = document.createElement('input');
    const file = new File(['logo'], 'logo.png', { type: 'image/png' });
    Object.defineProperty(input, 'files', { value: [file] });

    app.onLogoSelected({ target: input });
    await new Promise<void>((resolve, reject) => {
      const startedAt = Date.now();
      const waitForReader = () => {
        if (app.emailLogoData) {
          resolve();
        } else if (Date.now() - startedAt > 1000) {
          reject(new Error('El FileReader no terminó a tiempo.'));
        } else {
          setTimeout(waitForReader, 10);
        }
      };
      waitForReader();
    });

    expect(app.emailLogoName).toBe('logo.png');
    expect(app.emailLogoContentType).toBe('image/png');
    expect(app.emailLogoData).toBe('data:image/png;base64,bG9nbw==');
    expect(app.emailLogoPreview).toBe('data:image/png;base64,bG9nbw==');
  });

  it('should preserve a legacy logo URL when saving without a new logo', async () => {
    const fixture = TestBed.createComponent(App);
    const http = TestBed.inject(HttpTestingController);
    const app = fixture.componentInstance as any;
    app.emailTo = 'destino@example.com';
    app.emailSubject = 'Reporte';
    app.emailBodyHtml = '<p>Listo</p>';
    app.emailLogoData = '';
    app.emailLogoName = '';
    app.emailLogoContentType = '';
    app.emailLegacyLogoUrl = 'https://example.com/logo.png';

    const save = app.saveEmailConfig();
    const request = http.expectOne('/api/workflows/cajas/email-config');
    expect(request.request.body.logoUrl).toBe('https://example.com/logo.png');
    request.flush({
      config: {
        recipients: ['destino@example.com'],
        cc: [],
        bcc: [],
        subject: 'Reporte',
        bodyHtml: '<p>Listo</p>',
        logoUrl: 'https://example.com/logo.png',
        logoData: '',
        logoName: '',
        logoContentType: '',
      },
      message: 'Guardado',
    });
    await save;
    http.verify();
  });
});
