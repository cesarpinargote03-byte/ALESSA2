import { CommonModule, DatePipe, registerLocaleData } from '@angular/common';
import { HttpClient } from '@angular/common/http';
import { Component, ElementRef, HostListener, inject, OnInit, ViewChild } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { firstValueFrom } from 'rxjs';
import localeEs from '@angular/common/locales/es';
import { environment } from '../../environments/environment';

registerLocaleData(localeEs);

interface MailItem {
  id: string;
  subject: string;
  remitente_email: string;
  remitente_nombre?: string;
  categoria: string | null;
  categorias?: string[];
  ano_fiscal: number | null;
  carpeta_origen: string;
  received_at: string;
  semana_planificada_inicio?: string;
  semana_planificada_fin?: string;
  semana_iso?: string;
  semana_numero?: number;
  estado_planificacion?: string;
  outlook_event_id?: string;
  postergado_por_proyecto?: string;
}

interface ExcelPreviewRow {
  message_id: string;
  semana: string;
  postergado_por_proyecto: string;
  usuario: string;
  correo: string;
  asunto: string;
}

interface SentBatchRecord {
  batch_id: string;
  created_at: string;
  subject: string;
  recipients: string[];
  cc_recipients: string[];
  rows_count: number;
  filename: string;
  status: string;
  rows: ExcelPreviewRow[];
}

interface WeekOption {
  isoWeek: string;       // e.g. "2026-W44"
  weekNumber: number;    // e.g. 44
  year: number;          // e.g. 2026
  startDate: string;     // YYYY-MM-DD (Monday)
  endDate: string;       // YYYY-MM-DD (Friday)
  label: string;         // e.g. "Semana 44 · 26 oct – 30 oct 2026 (Esta semana)"
  shortRange: string;    // e.g. "26 oct – 30 oct"
}

type WorkspaceView = 'overview' | 'inbox' | 'planning' | 'batches' | 'settings';

const SAMPLE_MESSAGES: MailItem[] = [
  {
    id: 'demo-1',
    subject: 'Solicitud de revisión presupuestaria — sede norte',
    remitente_email: 'mariana.rojas@institucion.cl',
    categoria: 'Presupuesto',
    ano_fiscal: 2026,
    carpeta_origen: 'Bandeja de entrada',
    received_at: '2026-10-24T08:42:00-05:00',
  },
  {
    id: 'demo-2',
    subject: 'Actualización de antecedentes para convenio',
    remitente_email: 'pablo.munoz@institucion.cl',
    categoria: 'Convenios',
    ano_fiscal: 2026,
    carpeta_origen: 'Bandeja de entrada',
    received_at: '2026-10-24T08:16:00-05:00',
  },
  {
    id: 'demo-3',
    subject: 'Consulta sobre disponibilidad de fondos',
    remitente_email: 'camila.soto@institucion.cl',
    categoria: 'Finanzas',
    ano_fiscal: 2026,
    carpeta_origen: 'Solicitudes',
    received_at: '2026-10-23T16:35:00-05:00',
  },
  {
    id: 'demo-4',
    subject: 'Documentación complementaria proyecto regional',
    remitente_email: 'diego.fuentes@institucion.cl',
    categoria: 'Proyectos',
    ano_fiscal: 2026,
    carpeta_origen: 'Solicitudes',
    received_at: '2026-10-23T14:08:00-05:00',
  },
  {
    id: 'demo-5',
    subject: 'Re: Plan anual de adquisiciones',
    remitente_email: 'valentina.perez@institucion.cl',
    categoria: 'Adquisiciones',
    ano_fiscal: 2026,
    carpeta_origen: 'Bandeja de entrada',
    received_at: '2026-10-22T11:22:00-05:00',
  },
  {
    id: 'demo-6',
    subject: 'Solicitud de ajuste de cronograma',
    remitente_email: 'francisco.vera@institucion.cl',
    categoria: 'Planificación',
    ano_fiscal: 2025,
    carpeta_origen: 'Archivo',
    received_at: '2026-10-21T09:12:00-05:00',
  },
  {
    id: 'demo-7',
    subject: 'Informe de avance del segundo trimestre',
    remitente_email: 'sofia.contreras@institucion.cl',
    categoria: 'Informes',
    ano_fiscal: 2025,
    carpeta_origen: 'Bandeja de entrada',
    received_at: '2026-10-20T15:49:00-05:00',
  },
  {
    id: 'demo-8',
    subject: 'Solicitud de certificación presupuestaria',
    remitente_email: 'andres.silva@institucion.cl',
    categoria: 'Presupuesto',
    ano_fiscal: 2025,
    carpeta_origen: 'Solicitudes',
    received_at: '2026-10-20T10:03:00-05:00',
  },
  {
    id: 'demo-9',
    subject: 'Antecedentes para cierre de convenio',
    remitente_email: 'paula.torres@institucion.cl',
    categoria: 'Convenios',
    ano_fiscal: 2025,
    carpeta_origen: 'Archivo',
    received_at: '2026-10-19T13:28:00-05:00',
  },
];

const SHORT_MONTHS = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];

@Component({
  selector: 'app-home',
  templateUrl: 'home.page.html',
  styleUrls: ['home.page.scss'],
  imports: [CommonModule, DatePipe, FormsModule],
})
export class HomePage implements OnInit {
  activeView: WorkspaceView = 'inbox';
  messages: MailItem[] = environment.demoMode ? [...SAMPLE_MESSAGES] : [];
  authenticated = false;
  connectedEmail = '';
  usingDemoData = environment.demoMode;
  loading = false;
  syncing = false;
  errorMessage = '';
  successMessage = '';
  searchText = '';
  selectedCategory = '';
  selectedYear = '';
  selectedFolder = '';
  currentPage = 1;
  pageSize = 15;
  selectedMessage: MailItem | null = null;
  plannerMessage: MailItem | null = null;
  plannerWeek = '';
  plannerStart = '';
  plannerEnd = '';
  plannerTimezone = 'America/Guayaquil';
  plannerPostergado = 'No';
  planningBusy = false;
  availableWeeks: WeekOption[] = [];
  readonly maxPerWeek = 4;

  selectedPlanningWeek = '';
  excelPreviewOpen = false;
  excelPreviewRows: ExcelPreviewRow[] = [];
  excelRecipientsText = '';
  excelCcText = '';
  excelSubject = '';
  excelIntroText =
    'Estimados usuarios,\nAdjunto encontrarán el cuadro de planificación semanal de solicitudes registradas en Helpdesk.';
  excelFilename = 'Planificacion_Helpdesk.xlsx';
  excelSendingBusy = false;
  sentBatches: SentBatchRecord[] = [];
  darkMode = false;

  private readonly apiUrl = environment.apiUrl;
  private readonly localPlansKey = 'umbral_planned_messages_v1';

  @ViewChild('searchInput') private searchInput?: ElementRef<HTMLInputElement>;

  private readonly http = inject(HttpClient);

  constructor() {
    this.availableWeeks = this.buildUpcomingWeeks(20);
  }

  async ngOnInit(): Promise<void> {
    this.restoreLocalPlans();
    try {
      const status = await firstValueFrom(
        this.http.get<{ authenticated: boolean }>(`${this.apiUrl}/auth/status`, {
          withCredentials: true,
        }),
      );
      this.authenticated = status.authenticated;
      if (this.authenticated) {
        await this.loadConnectedUser();
        await this.loadMessages();
        if (!this.messages.length) {
          await this.syncInbox();
        }
      } else if (!environment.demoMode) {
        this.messages = [];
        this.usingDemoData = false;
      }
    } catch {
      if (!environment.demoMode) {
        this.errorMessage = 'No se pudo conectar con el servidor. Comprueba que el backend esté activo.';
      }
    }
  }

  @HostListener('window:keydown', ['$event'])
  handleKeyboardShortcut(event: KeyboardEvent): void {
    if ((event.ctrlKey || event.metaKey) && event.key.toLocaleLowerCase() === 'k') {
      event.preventDefault();
      this.searchInput?.nativeElement.focus();
    }
  }

  get filteredMessages(): MailItem[] {
    const search = this.searchText.trim().toLocaleLowerCase('es');
    return this.messages.filter((message) => {
      const matchesSearch =
        !search ||
        message.subject.toLocaleLowerCase('es').includes(search) ||
        message.remitente_email.toLocaleLowerCase('es').includes(search) ||
        (message.remitente_nombre ?? '').toLocaleLowerCase('es').includes(search);
      const matchesCategory =
        !this.selectedCategory ||
        message.categoria === this.selectedCategory ||
        (message.categorias ?? []).includes(this.selectedCategory);
      const matchesYear = !this.selectedYear || String(message.ano_fiscal) === this.selectedYear;
      const matchesFolder = !this.selectedFolder || message.carpeta_origen === this.selectedFolder;
      return matchesSearch && matchesCategory && matchesYear && matchesFolder;
    });
  }

  get plannedMessages(): MailItem[] {
    return this.messages
      .filter((m) => !!m.semana_planificada_inicio)
      .sort((a, b) => (a.semana_planificada_inicio ?? '').localeCompare(b.semana_planificada_inicio ?? ''));
  }

  get filteredPlannedMessages(): MailItem[] {
    if (!this.selectedPlanningWeek) {
      return this.plannedMessages;
    }
    return this.plannedMessages.filter((m) => m.semana_iso === this.selectedPlanningWeek);
  }

  get plannedWeeksSummary(): { isoWeek: string; label: string; count: number }[] {
    const counts = new Map<string, { label: string; count: number }>();
    for (const msg of this.plannedMessages) {
      const iso = msg.semana_iso || 'Sin semana';
      const existing = counts.get(iso);
      if (existing) {
        existing.count += 1;
      } else {
        counts.set(iso, {
          label: this.formatPlanBadge(msg),
          count: 1,
        });
      }
    }
    return Array.from(counts.entries()).map(([isoWeek, info]) => ({
      isoWeek,
      label: info.label,
      count: info.count,
    }));
  }

  getWeekCount(isoWeek: string, excludeMessageId?: string): number {
    if (!isoWeek) {
      return 0;
    }
    return this.plannedMessages.filter(
      (m) => m.semana_iso === isoWeek && m.id !== excludeMessageId,
    ).length;
  }

  formatWeekOptionLabel(wk: WeekOption): string {
    const count = this.getWeekCount(wk.isoWeek, this.plannerMessage?.id);
    const status = count >= this.maxPerWeek ? `LLENA (${count}/${this.maxPerWeek})` : `${count}/${this.maxPerWeek} cupos`;
    return `${wk.label} · [${status}]`;
  }

  get pageMessages(): MailItem[] {
    const start = (this.currentPage - 1) * this.pageSize;
    return this.filteredMessages.slice(start, start + this.pageSize);
  }

  get totalPages(): number {
    return Math.max(1, Math.ceil(this.filteredMessages.length / this.pageSize));
  }

  get categories(): string[] {
    const allCats: string[] = [];
    for (const message of this.messages) {
      if (message.categorias?.length) {
        allCats.push(...message.categorias);
      } else if (message.categoria) {
        allCats.push(message.categoria);
      }
    }
    return [...new Set(allCats.filter((value): value is string => !!value))].sort((a, b) =>
      a.localeCompare(b, 'es'),
    );
  }

  get years(): number[] {
    return [...new Set(this.messages.map((message) => message.ano_fiscal).filter((value): value is number => value !== null))]
      .sort((a, b) => b - a);
  }

  get folders(): string[] {
    return [...new Set(this.messages.map((message) => message.carpeta_origen))].sort((a, b) =>
      a.localeCompare(b, 'es'),
    );
  }

  get startItem(): number {
    return this.filteredMessages.length ? (this.currentPage - 1) * this.pageSize + 1 : 0;
  }

  get endItem(): number {
    return Math.min(this.currentPage * this.pageSize, this.filteredMessages.length);
  }

  get unreadCount(): number {
    return this.messages.length ? Math.min(4, Math.ceil(this.messages.length / 3)) : 0;
  }

  get failureCount(): number {
    return this.messages.filter((m) => this.getMessageSemanticStatus(m) === 'failure').length;
  }

  get successCount(): number {
    return this.messages.filter((m) => this.getMessageSemanticStatus(m) === 'success').length;
  }

  getMessageSemanticStatus(message: MailItem): 'failure' | 'success' | 'neutral' {
    const subj = (message.subject || '').toLocaleLowerCase('es');
    if (
      subj.includes('failure') ||
      subj.includes('fallo') ||
      subj.includes('error') ||
      subj.includes('caída') ||
      subj.includes('caida') ||
      subj.includes('cierre inesperado') ||
      subj.includes('critical') ||
      subj.includes('down')
    ) {
      return 'failure';
    }
    if (
      subj.includes('success') ||
      subj.includes('restablecido') ||
      subj.includes('recovered') ||
      subj.includes('resuelto')
    ) {
      return 'success';
    }
    return 'neutral';
  }

  toggleDarkMode(): void {
    this.darkMode = !this.darkMode;
    document.body.classList.toggle('dark', this.darkMode);
  }

  setView(view: WorkspaceView): void {
    this.activeView = view;
    this.errorMessage = '';
    this.successMessage = '';
  }

  resetFilters(): void {
    this.searchText = '';
    this.selectedCategory = '';
    this.selectedYear = '';
    this.selectedFolder = '';
    this.currentPage = 1;
  }

  changePage(page: number): void {
    if (page >= 1 && page <= this.totalPages) {
      this.currentPage = page;
    }
  }

  async syncInbox(): Promise<void> {
    this.errorMessage = '';
    this.successMessage = '';
    if (!this.authenticated) {
      this.errorMessage = 'Conecta tu cuenta Microsoft para sincronizar la bandeja real.';
      return;
    }

    this.syncing = true;
    try {
      const syncResult = await firstValueFrom(
        this.http.post<{ count: number; messages: MailItem[] }>(
          `${this.apiUrl}/outlook/sync?limit=2000`,
          {},
          { withCredentials: true },
        ),
      );
      this.messages = syncResult.messages ?? [];
      this.restoreLocalPlans();
      this.usingDemoData = false;
      this.currentPage = 1;
      this.successMessage = `Bandejas sincronizadas (${this.messages.length} correos en ${this.folders.length} carpetas).`;
    } catch (error: unknown) {
      this.errorMessage = this.getApiError(error, 'No se pudo sincronizar Outlook.');
    } finally {
      this.syncing = false;
    }
  }

  async startLogin(): Promise<void> {
    this.errorMessage = '';
    this.successMessage = '';
    try {
      const loginResp = await firstValueFrom(
        this.http.post<{ authenticated: boolean; email?: string }>(
          `${this.apiUrl}/auth/login-local`,
          {},
          { withCredentials: true },
        ),
      );
      this.authenticated = true;
      this.connectedEmail = loginResp.email ?? '';
      if (!this.connectedEmail) {
        await this.loadConnectedUser();
      }
      await this.syncInbox();
      if (this.connectedEmail) {
        this.successMessage = `Conectado a Outlook como ${this.connectedEmail} (${this.messages.length} correos sincronizados).`;
      }
    } catch {
      window.location.assign(`${this.apiUrl}/auth/login`);
    }
  }

  private async loadConnectedUser(): Promise<void> {
    try {
      const me = await firstValueFrom(
        this.http.get<{ email: string }>(`${this.apiUrl}/auth/me`, {
          withCredentials: true,
        }),
      );
      this.connectedEmail = me.email ?? '';
    } catch {
      // ignore if /auth/me fails
    }
  }

  openPlanner(message: MailItem, event: Event): void {
    event.stopPropagation();
    this.selectedMessage = null;
    this.plannerMessage = message;
    this.plannerPostergado = message.postergado_por_proyecto || 'No';
    this.errorMessage = '';

    let defaultWeek = message.semana_iso || '';
    if (!defaultWeek) {
      const firstAvailable = this.availableWeeks.find(
        (wk) => this.getWeekCount(wk.isoWeek, message.id) < this.maxPerWeek,
      );
      defaultWeek = firstAvailable?.isoWeek || this.availableWeeks[0]?.isoWeek || '';
    }
    this.onWeekChange(defaultWeek);
  }

  onWeekChange(isoWeekValue: string): void {
    this.plannerWeek = isoWeekValue;
    const preset = this.availableWeeks.find((w) => w.isoWeek === isoWeekValue);
    if (preset) {
      this.plannerStart = preset.startDate;
      this.plannerEnd = preset.endDate;
      return;
    }
    const computed = this.computeWeekRangeFromIso(isoWeekValue);
    if (computed) {
      this.plannerStart = computed.startDate;
      this.plannerEnd = computed.endDate;
    }
  }

  formatPlanBadge(message: MailItem): string {
    if (!message.semana_planificada_inicio) {
      return '';
    }
    const start = this.formatShortDate(message.semana_planificada_inicio);
    const end = message.semana_planificada_fin ? this.formatShortDate(message.semana_planificada_fin) : '';
    const weekLabel = message.semana_numero ? `Sem. ${message.semana_numero}` : 'Planificado';
    return end ? `${weekLabel} · ${start} – ${end}` : `${weekLabel} · ${start}`;
  }

  formatShortDate(isoDate: string): string {
    const parts = isoDate.slice(0, 10).split('-').map(Number);
    if (parts.length !== 3 || parts.some(isNaN)) {
      return isoDate;
    }
    const [, month, day] = parts;
    return `${day} ${SHORT_MONTHS[(month - 1) % 12]}`;
  }

  async savePlan(): Promise<void> {
    if (!this.plannerMessage || !this.plannerStart || !this.plannerEnd) {
      this.errorMessage = 'Selecciona una semana para planificar la solicitud.';
      return;
    }
    if (this.plannerEnd < this.plannerStart) {
      this.errorMessage = 'La fecha de término debe ser posterior al inicio.';
      return;
    }

    const currentInWeek = this.getWeekCount(this.plannerWeek, this.plannerMessage.id);
    if (currentInWeek >= this.maxPerWeek) {
      const weekNum = this.extractWeekNumber(this.plannerWeek) || this.plannerWeek;
      this.errorMessage = `La Semana ${weekNum} ya tiene el máximo de ${this.maxPerWeek} solicitudes planificadas. Elige la siguiente semana disponible.`;
      return;
    }

    if (!this.authenticated) {
      this.errorMessage = 'Conecta tu cuenta de Outlook antes de planificar.';
      return;
    }

    this.planningBusy = true;
    this.errorMessage = '';
    try {
      const response = await firstValueFrom(
        this.http.post<{
          message_id: string;
          event_id: string;
          status: string;
          start_date?: string;
          end_date?: string;
          semana_iso?: string;
          semana_numero?: number;
          postergado_por_proyecto?: string;
        }>(
          `${this.apiUrl}/planificacion/asignar`,
          {
            message_id: this.plannerMessage.id,
            start_date: this.plannerStart,
            end_date: this.plannerEnd,
            time_zone: this.plannerTimezone,
            subject: this.plannerMessage.subject,
            postergado_por_proyecto: this.plannerPostergado || 'No',
          },
          { withCredentials: true },
        ),
      );

      const weekNum = response.semana_numero ?? this.extractWeekNumber(this.plannerWeek);
      const weekIso = response.semana_iso ?? this.plannerWeek;

      this.applyPlanToMessage(this.plannerMessage.id, {
        semana_planificada_inicio: this.plannerStart,
        semana_planificada_fin: this.plannerEnd,
        semana_iso: weekIso,
        semana_numero: weekNum,
        estado_planificacion: 'Planificado',
        outlook_event_id: response.event_id,
        postergado_por_proyecto: (this.plannerPostergado || 'No').trim() || 'No',
      });

      const newCount = this.getWeekCount(weekIso);
      const rangeText = `${this.formatShortDate(this.plannerStart)} – ${this.formatShortDate(this.plannerEnd)}`;
      this.successMessage = `Solicitud planificada en Semana ${weekNum || ''} (${rangeText}) · Cupo semanal: ${newCount}/${this.maxPerWeek}.`;
      this.plannerMessage = null;
    } catch (error: unknown) {
      this.errorMessage = this.getApiError(error, 'No se pudo crear el evento en Outlook.');
    } finally {
      this.planningBusy = false;
    }
  }

  unplanMessage(message: MailItem, event: Event): void {
    event.stopPropagation();
    this.messages = this.messages.map((item) => {
      if (item.id !== message.id) {
        return item;
      }
      return {
        ...item,
        semana_planificada_inicio: undefined,
        semana_planificada_fin: undefined,
        semana_iso: undefined,
        semana_numero: undefined,
        estado_planificacion: undefined,
        outlook_event_id: undefined,
        postergado_por_proyecto: undefined,
      };
    });
    try {
      const storedRaw = localStorage.getItem(this.localPlansKey);
      if (storedRaw) {
        const stored: Record<string, Partial<MailItem>> = JSON.parse(storedRaw);
        delete stored[message.id];
        localStorage.setItem(this.localPlansKey, JSON.stringify(stored));
      }
    } catch {
      // ignore storage errors
    }
    this.successMessage = 'Solicitud retirada de la planificación semanal.';
  }

  openExcelPreview(): void {
    this.errorMessage = '';
    const sourceItems = this.filteredPlannedMessages;
    if (!sourceItems.length) {
      this.errorMessage = 'No hay solicitudes planificadas para generar el formato de Excel.';
      return;
    }

    this.excelPreviewRows = sourceItems.map((item) => {
      const userDisplay =
        item.remitente_nombre?.trim() ||
        item.remitente_email.split('@')[0].replace(/[._]/g, ' ');
      return {
        message_id: item.id,
        semana: this.formatPlanBadge(item),
        postergado_por_proyecto: item.postergado_por_proyecto || 'No',
        usuario: userDisplay,
        correo: item.remitente_email,
        asunto: item.subject,
      };
    });

    const uniqueEmails = [
      ...new Set(
        sourceItems
          .map((item) => item.remitente_email.trim())
          .filter((email) => email.includes('@')),
      ),
    ];
    this.excelRecipientsText = uniqueEmails.join('; ');
    if (this.connectedEmail && !this.excelCcText) {
      this.excelCcText = this.connectedEmail;
    }

    const weekTag = this.selectedPlanningWeek
      ? `Semana ${this.extractWeekNumber(this.selectedPlanningWeek) || this.selectedPlanningWeek}`
      : 'Consolidado Semanal';
    this.excelSubject = `Planificación Semanal de Solicitudes (${weekTag}) - Helpdesk`;
    this.excelFilename = `Planificacion_Helpdesk_${weekTag.replace(/\s+/g, '_')}.xlsx`;
    this.excelPreviewOpen = true;
  }

  async downloadExcelFile(): Promise<void> {
    if (!this.excelPreviewRows.length) {
      this.openExcelPreview();
    }
    if (!this.excelPreviewRows.length) {
      return;
    }
    try {
      const blob = await firstValueFrom(
        this.http.post(
          `${this.apiUrl}/planificacion/excel`,
          {
            filename: this.excelFilename,
            rows: this.excelPreviewRows,
          },
          { responseType: 'blob', withCredentials: true },
        ),
      );
      const url = window.URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = this.excelFilename || 'Planificacion_Helpdesk.xlsx';
      anchor.click();
      window.URL.revokeObjectURL(url);
      this.successMessage = `Archivo Excel "${this.excelFilename}" descargado correctamente.`;
    } catch (error: unknown) {
      this.errorMessage = this.getApiError(error, 'No se pudo generar el archivo Excel.');
    }
  }

  async approveAndSendExcel(mode: 'send' | 'display'): Promise<void> {
    const recipients = this.excelRecipientsText
      .split(/[;,]/)
      .map((r) => r.trim())
      .filter((r) => r.includes('@'));
    const ccRecipients = this.excelCcText
      .split(/[;,]/)
      .map((r) => r.trim())
      .filter((r) => r.includes('@'));

    if (!recipients.length) {
      this.errorMessage = 'Ingresa al menos un correo destinatario válido para aprobar el envío.';
      return;
    }

    this.excelSendingBusy = true;
    this.errorMessage = '';
    try {
      const batch = await firstValueFrom(
        this.http.post<SentBatchRecord>(
          `${this.apiUrl}/planificacion/enviar-excel`,
          {
            recipients,
            cc_recipients: ccRecipients,
            subject: this.excelSubject,
            intro_text: this.excelIntroText,
            filename: this.excelFilename,
            mode,
            rows: this.excelPreviewRows,
          },
          { withCredentials: true },
        ),
      );
      this.sentBatches = [batch, ...this.sentBatches];
      this.excelPreviewOpen = false;
      if (mode === 'display') {
        this.successMessage =
          'Se abrió el borrador en tu Outlook con el Excel adjunto y la tabla lista para tu revisión final.';
      } else {
        this.successMessage = `Planificación aprobada y enviada por Outlook a ${recipients.length} destinatario(s) con el archivo ${this.excelFilename} adjunto.`;
      }
    } catch (error: unknown) {
      this.errorMessage = this.getApiError(
        error,
        'No se pudo enviar la planificación desde Outlook.',
      );
    } finally {
      this.excelSendingBusy = false;
    }
  }

  clearNotices(): void {
    this.errorMessage = '';
    this.successMessage = '';
  }

  private applyPlanToMessage(messageId: string, plan: Partial<MailItem>): void {
    this.messages = this.messages.map((item) =>
      item.id === messageId ? { ...item, ...plan } : item,
    );
    try {
      const storedRaw = localStorage.getItem(this.localPlansKey);
      const stored: Record<string, Partial<MailItem>> = storedRaw ? JSON.parse(storedRaw) : {};
      stored[messageId] = plan;
      localStorage.setItem(this.localPlansKey, JSON.stringify(stored));
    } catch {
      // ignore storage errors
    }
  }

  private restoreLocalPlans(): void {
    try {
      const storedRaw = localStorage.getItem(this.localPlansKey);
      if (!storedRaw) {
        return;
      }
      const stored: Record<string, Partial<MailItem>> = JSON.parse(storedRaw);
      this.messages = this.messages.map((item) =>
        stored[item.id] ? { ...item, ...stored[item.id] } : item,
      );
    } catch {
      // ignore storage errors
    }
  }

  private extractWeekNumber(isoWeek: string): number | undefined {
    const match = /W(\d{1,2})$/i.exec(isoWeek);
    return match ? Number(match[1]) : undefined;
  }

  private buildUpcomingWeeks(count: number): WeekOption[] {
    const today = new Date();
    const dayOfWeek = today.getDay(); // 0=Sun, 1=Mon, ...
    const diffToMonday = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
    const baseMonday = new Date(today.getFullYear(), today.getMonth(), today.getDate() + diffToMonday);

    const options: WeekOption[] = [];
    for (let i = 0; i < count; i++) {
      const monday = new Date(baseMonday.getFullYear(), baseMonday.getMonth(), baseMonday.getDate() + i * 7);
      const friday = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + 4);
      const { year, week } = this.getIsoWeek(monday);
      const isoWeek = `${year}-W${String(week).padStart(2, '0')}`;
      const startDate = this.toIsoDateString(monday);
      const endDate = this.toIsoDateString(friday);
      const shortRange = `${monday.getDate()} ${SHORT_MONTHS[monday.getMonth()]} – ${friday.getDate()} ${SHORT_MONTHS[friday.getMonth()]}`;
      const tag = i === 0 ? ' · Esta semana' : i === 1 ? ' · Próxima semana' : '';
      options.push({
        isoWeek,
        weekNumber: week,
        year,
        startDate,
        endDate,
        shortRange,
        label: `Semana ${week} (${shortRange} ${friday.getFullYear()})${tag}`,
      });
    }
    return options;
  }

  private computeWeekRangeFromIso(isoWeek: string): { startDate: string; endDate: string } | null {
    const match = /^(\d{4})-W(\d{1,2})$/i.exec(isoWeek.trim());
    if (!match) {
      return null;
    }
    const year = Number(match[1]);
    const week = Number(match[2]);
    const jan4 = new Date(year, 0, 4);
    const jan4Day = jan4.getDay() || 7;
    const mondayWeek1 = new Date(year, 0, 4 - (jan4Day - 1));
    const monday = new Date(mondayWeek1.getFullYear(), mondayWeek1.getMonth(), mondayWeek1.getDate() + (week - 1) * 7);
    const friday = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + 4);
    return {
      startDate: this.toIsoDateString(monday),
      endDate: this.toIsoDateString(friday),
    };
  }

  private getIsoWeek(d: Date): { year: number; week: number } {
    const target = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    const dayNr = (target.getDay() + 6) % 7;
    target.setDate(target.getDate() - dayNr + 3);
    const firstThursday = new Date(target.getFullYear(), 0, 4);
    const diff = target.getTime() - firstThursday.getTime();
    const week = 1 + Math.round(diff / (7 * 24 * 3600 * 1000));
    return { year: target.getFullYear(), week };
  }

  private toIsoDateString(d: Date): string {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  private async loadMessages(): Promise<void> {
    this.loading = true;
    try {
      const result = await firstValueFrom(
        this.http.get<{ items: MailItem[] }>(`${this.apiUrl}/correos?page_size=10000`, {
          withCredentials: true,
        }),
      );
      this.messages = result.items ?? [];
      this.restoreLocalPlans();
      this.usingDemoData = false;
      this.currentPage = 1;
    } catch (error: unknown) {
      this.errorMessage = this.getApiError(error, 'No se pudieron cargar los mensajes.');
    } finally {
      this.loading = false;
    }
  }

  private getApiError(error: unknown, fallback: string): string {
    if (typeof error === 'object' && error !== null && 'error' in error) {
      const payload = (error as { error?: { detail?: unknown } }).error;
      if (typeof payload?.detail === 'string') {
        return payload.detail;
      }
    }
    return fallback;
  }
}
