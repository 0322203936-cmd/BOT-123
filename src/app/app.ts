import { CommonModule } from '@angular/common';
import { HttpClient, HttpErrorResponse } from '@angular/common/http';
import { Component, computed, OnDestroy, OnInit, signal } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { firstValueFrom } from 'rxjs';

type WorkflowKey = 'galleria' | 'cancelaciones' | 'pegarData' | 'inventario' | 'reunion';
type RunState = 'idle' | 'queued' | 'in_progress' | 'completed' | 'unknown';

interface WorkflowRun {
  id: number;
  status: RunState;
  conclusion: string | null;
  createdAt: string;
  updatedAt: string;
  url: string;
  event: string;
}

interface WorkflowStatus {
  key: WorkflowKey;
  name: string;
  description: string;
  schedule: string;
  run: WorkflowRun | null;
}

interface ApiConfig {
  authRequired: boolean;
  configured: boolean;
}

@Component({
  selector: 'app-root',
  imports: [CommonModule, FormsModule],
  templateUrl: './app.html',
  styleUrl: './app.css',
})
export class App implements OnInit, OnDestroy {
  protected readonly workflows = signal<WorkflowStatus[]>([]);
  protected readonly visibleWorkflows = computed(() =>
    this.workflows().filter((workflow) => workflow.key !== 'cancelaciones'),
  );
  protected readonly loading = signal(true);
  protected readonly running = signal<WorkflowKey | null>(null);
  protected readonly message = signal('');
  protected readonly error = signal('');
  protected readonly authRequired = signal(false);
  protected readonly authenticated = signal(false);
  protected readonly configured = signal(true);
  protected readonly clock = signal(Date.now());
  protected password = '';

  private refreshTimer?: ReturnType<typeof setInterval>;
  private clockTimer?: ReturnType<typeof setInterval>;

  constructor(private readonly http: HttpClient) {}

  async ngOnInit(): Promise<void> {
    const savedPassword = sessionStorage.getItem('atajos-password') ?? '';
    this.password = savedPassword;

    try {
      const config = await firstValueFrom(this.http.get<ApiConfig>('/api/config'));
      this.authRequired.set(config.authRequired);
      this.configured.set(config.configured);

      if (!config.authRequired || savedPassword) {
        await this.connect(false);
      } else {
        this.loading.set(false);
      }
    } catch {
      this.loading.set(false);
      this.error.set('No fue posible conectar con el servidor. Intenta nuevamente.');
    }

    this.refreshTimer = setInterval(() => {
      if (this.authenticated()) void this.refresh(false);
    }, 15_000);
    this.clockTimer = setInterval(() => this.clock.set(Date.now()), 1_000);
  }

  ngOnDestroy(): void {
    if (this.refreshTimer) clearInterval(this.refreshTimer);
    if (this.clockTimer) clearInterval(this.clockTimer);
  }

  protected async connect(showError = true): Promise<void> {
    this.error.set('');
    if (this.password) sessionStorage.setItem('atajos-password', this.password);

    try {
      await this.refresh(false);
      this.authenticated.set(true);
    } catch (error) {
      this.authenticated.set(false);
      sessionStorage.removeItem('atajos-password');
      if (showError) this.error.set(this.getErrorMessage(error));
    } finally {
      this.loading.set(false);
    }
  }

  protected async refresh(showError = true): Promise<void> {
    try {
      const response = await firstValueFrom(
        this.http.get<{ workflows: WorkflowStatus[] }>('/api/workflows', {
          headers: this.authHeaders(),
        }),
      );
      this.workflows.set(response.workflows);
      this.authenticated.set(true);
      if (showError) this.error.set('');
    } catch (error) {
      if (showError) this.error.set(this.getErrorMessage(error));
      throw error;
    }
  }

  protected async execute(workflow: WorkflowStatus): Promise<void> {
    if (this.running() || this.isActive(workflow.run)) return;
    if (!this.confirmProtectedExecution(workflow)) {
      return;
    }

    this.running.set(workflow.key);
    this.error.set('');
    this.message.set('');

    try {
      const response = await firstValueFrom(
        this.http.post<{ message: string; dispatchedAt: string }>(
          `/api/workflows/${workflow.key}/dispatch`,
          {},
          { headers: this.authHeaders() },
        ),
      );
      this.workflows.update((items) =>
        items.map((item) =>
          item.key === workflow.key
            ? {
                ...item,
                run: {
                  id: 0,
                  status: 'queued',
                  conclusion: null,
                  createdAt: response.dispatchedAt,
                  updatedAt: response.dispatchedAt,
                  url: '',
                  event: 'workflow_dispatch',
                },
              }
            : item,
        ),
      );
      this.message.set(response.message);
      await new Promise((resolve) => setTimeout(resolve, 2_000));
      await this.refresh(false);
    } catch (error) {
      this.error.set(this.getErrorMessage(error));
    } finally {
      this.running.set(null);
    }
  }

  private confirmProtectedExecution(workflow: WorkflowStatus): boolean {
    if (!(['pegarData', 'inventario', 'reunion'] as WorkflowKey[]).includes(workflow.key)) {
      return true;
    }

    const response = window.prompt(
      `Para ejecutar ${workflow.name}, escribe la palabra CONFIRMAR.`,
    );
    if ((response ?? '').trim().toUpperCase() === 'CONFIRMAR') {
      return true;
    }

    this.message.set('');
    this.error.set('Ejecución cancelada: debes escribir CONFIRMAR.');
    return false;
  }

  protected isActive(run: WorkflowRun | null): boolean {
    return run?.status === 'queued' || run?.status === 'in_progress';
  }

  protected statusLabel(run: WorkflowRun | null): string {
    if (!run) return 'Sin ejecuciones';
    if (run.status === 'queued') return 'En cola';
    if (run.status === 'in_progress') return 'Ejecutándose';
    if (run.conclusion === 'success') return 'Listo';
    if (run.conclusion === 'failure') return 'Falló';
    if (run.conclusion === 'cancelled') return 'Cancelado';
    return 'Finalizado';
  }

  protected statusClass(run: WorkflowRun | null): string {
    if (!run) return 'neutral';
    if (this.isActive(run)) return 'active';
    if (run.conclusion === 'success') return 'success';
    if (run.conclusion === 'failure' || run.conclusion === 'cancelled') return 'failure';
    return 'neutral';
  }

  protected formatDate(value: string | undefined): string {
    if (!value) return 'Todavía no hay información';
    return new Intl.DateTimeFormat('es-MX', {
      dateStyle: 'medium',
      timeStyle: 'short',
    }).format(new Date(value));
  }

  protected elapsedLabel(run: WorkflowRun | null): string {
    if (!run?.createdAt) return '00:00';
    const totalSeconds = Math.max(0, Math.floor((this.clock() - new Date(run.createdAt).getTime()) / 1_000));
    const hours = Math.floor(totalSeconds / 3_600);
    const minutes = Math.floor((totalSeconds % 3_600) / 60);
    const seconds = totalSeconds % 60;
    const mm = String(minutes).padStart(2, '0');
    const ss = String(seconds).padStart(2, '0');
    return hours > 0 ? `${String(hours).padStart(2, '0')}:${mm}:${ss}` : `${mm}:${ss}`;
  }

  private authHeaders(): Record<string, string> {
    return this.password ? { 'X-App-Password': this.password } : {};
  }

  private getErrorMessage(error: unknown): string {
    if (error instanceof HttpErrorResponse) {
      if (error.status === 401) return 'La contraseña no es correcta.';
      if (typeof error.error?.message === 'string') return error.error.message;
    }
    return 'Ocurrió un error inesperado. Intenta nuevamente.';
  }
}
