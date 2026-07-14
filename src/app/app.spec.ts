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
});
