import { Component, signal, ChangeDetectorRef } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import * as XLSX from 'xlsx';
import { ApiService } from '../services/api.service';

interface MapeoColumnas {
  numero_inventario: string;
  descripcion: string;
  clasificacion: string;
  ubicacion_esperada: string;
}

@Component({
  selector: 'app-importador-catalogo',
  standalone: true,
  imports: [CommonModule, FormsModule],
  template: `
<button (click)="abrir()"
        class="btn-ghost flex items-center gap-1.5 text-sm">
      <svg class="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2"
          d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5"/>
      </svg>
      Importar catálogo
    </button>

    @if (abierto()) {
      <div class="fixed inset-0 bg-guinda-950/60 backdrop-blur-sm z-50
                  flex items-center justify-center p-4"
           (click)="cerrar()">
        <div class="bg-white rounded-2xl shadow-modal w-full max-w-4xl
                    max-h-[90vh] flex flex-col animate-slide-up overflow-hidden"
             (click)="$event.stopPropagation()">

          <div class="h-1.5 bg-gradient-to-r from-guinda-700 via-dorado to-guinda-700 shrink-0"></div>
          <div class="px-6 py-4 border-b border-crema-border flex items-center justify-between shrink-0">
            <div>
              <h3 class="font-display text-xl text-guinda-800 font-semibold">
                Importar catálogo de bienes
              </h3>
              <p class="text-guinda-400 text-xs mt-0.5">
                Sube cualquier Excel — mapea las columnas y el sistema hace el resto
              </p>
            </div>
            <button (click)="cerrar()" class="text-guinda-300 hover:text-guinda-600 transition-colors">
              <svg class="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/>
              </svg>
            </button>
          </div>

          <div class="px-6 pt-4 pb-0 shrink-0">
            <div class="flex items-center gap-2 mb-5">
              @for (p of [1,2,3]; track p) {
                <div class="flex items-center gap-2">
                  <div class="w-7 h-7 rounded-full flex items-center justify-center text-xs font-bold
                               transition-all duration-300"
                       [class]="paso() >= p
                         ? 'bg-guinda-700 text-white'
                         : 'bg-crema-dark text-guinda-400'">
                    @if (paso() > p) {
                      <svg class="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="3" d="M4.5 12.75l6 6 9-13.5"/>
                      </svg>
                    } @else {
                      {{ p }}
                    }
                  </div>
                  <span class="text-xs hidden sm:block"
                        [class]="paso() >= p ? 'text-guinda-700 font-medium' : 'text-guinda-300'">
                    {{ p === 1 ? 'Subir archivo' : p === 2 ? 'Mapear columnas' : 'Confirmar' }}
                  </span>
                  @if (p < 3) {
                    <div class="w-8 h-px mx-1" [class]="paso() > p ? 'bg-guinda-500' : 'bg-crema-border'"></div>
                  }
                </div>
              }
            </div>
          </div>

          <div class="flex-1 overflow-y-auto px-6 pb-6 space-y-5">
            @if (paso() === 1) {
              <div class="border-2 border-dashed border-crema-border rounded-xl p-10
                          text-center hover:border-guinda-300 hover:bg-guinda-50
                          transition-all duration-200 cursor-pointer"
                   (click)="inputArchivo.click()"
                   (dragover)="$event.preventDefault()"
                   (drop)="onDrop($event)">
                <input #inputArchivo type="file" accept=".xlsx,.xls,.csv"
                       class="hidden" (change)="onArchivoSeleccionado($event)"/>

                @if (!archivo()) {
                  <svg class="w-14 h-14 text-guinda-200 mx-auto mb-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5"
                      d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0
                         011.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m6.75 12l-3-3m0
                         0l-3 3m3-3v6m-1.5-15H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0
                         .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504
                         1.125-1.125V11.25a9 9 0 00-9-9z"/>
                  </svg>
                  <p class="text-guinda-600 font-semibold text-lg mb-1">
                    Arrastra tu Excel aquí
                  </p>
                  <p class="text-guinda-400 text-sm mb-3">o haz clic para seleccionar</p>
                  <span class="text-xs bg-crema border border-crema-border text-guinda-400
                               px-3 py-1 rounded-full">.xlsx · .xls · .csv</span>
                } @else {
                  <svg class="w-12 h-12 text-success-DEFAULT mx-auto mb-3" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2"
                      d="M9 12.75L11.25 15 15 9.75M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/>
                  </svg>
                  <p class="text-guinda-800 font-semibold text-lg">{{ archivo()!.name }}</p>
                  <p class="text-guinda-400 text-sm mt-1 mb-3">
                    <span class="font-mono font-bold text-guinda-600">{{ filasPrevia().length }}</span>
                    filas detectadas ·
                    <span class="font-mono font-bold text-guinda-600">{{ columnasExcel().length }}</span>
                    columnas
                  </p>
                  <button (click)="$event.stopPropagation(); resetArchivo()"
                          class="text-xs text-guinda-400 hover:text-danger-DEFAULT underline transition-colors">
                    Cambiar archivo
                  </button>
                }
              </div>

              @if (errorArchivo()) {
                <div class="flex items-center gap-2 bg-danger-bg border border-danger-border
                            rounded-lg px-4 py-3 animate-fade-in">
                  <svg class="w-4 h-4 text-danger-DEFAULT shrink-0" fill="currentColor" viewBox="0 0 20 20">
                    <path fill-rule="evenodd"
                      d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-8-5a.75.75 0 01.75.75v4.5a.75.75 0 01-1.5 0v-4.5A.75.75 0 0110 5zm0 10a1 1 0 100-2 1 1 0 000 2z"
                      clip-rule="evenodd"/>
                  </svg>
                  <span class="text-sm text-danger-DEFAULT">{{ errorArchivo() }}</span>
                </div>
              }

              @if (archivo() && filasPrevia().length > 0) {
                <div class="flex justify-end">
                  <button (click)="paso.set(2)" class="btn-primary flex items-center gap-2">
                    Continuar
                    <svg class="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13.5 4.5L21 12m0 0l-7.5 7.5M21 12H3"/>
                    </svg>
                  </button>
                </div>
              }
            }

            @if (paso() === 2) {
              <div>
                <p class="text-xs font-semibold text-guinda-500 uppercase tracking-wider mb-2">
                  Vista previa de tu archivo (primeras 4 filas)
                </p>
                <div class="overflow-x-auto rounded-lg border border-crema-border">
                  <table class="text-xs w-full">
                    <thead class="bg-guinda-50">
                      <tr>
                        @for (col of columnasExcel(); track col) {
                          <th class="px-3 py-2 text-left font-semibold text-guinda-600
                                     border-b border-crema-border whitespace-nowrap">
                            {{ col }}
                          </th>
                        }
                      </tr>
                    </thead>
                    <tbody>
                      @for (fila of filasPrevia().slice(0, 4); track $index) {
                        <tr class="border-b border-crema-border last:border-0 hover:bg-crema/50">
                          @for (col of columnasExcel(); track col) {
                            <td class="px-3 py-2 text-guinda-500 whitespace-nowrap max-w-36 truncate">
                              {{ fila[col] ?? '—' }}
                            </td>
                          }
                        </tr>
                      }
                    </tbody>
                  </table>
                </div>
              </div>

              <div>
                <p class="text-xs font-semibold text-guinda-500 uppercase tracking-wider mb-3">
                  ¿Qué columna corresponde a cada campo?
                </p>
                <div class="grid sm:grid-cols-2 gap-4">
                  <div>
                    <label class="block text-xs font-semibold text-guinda-600 uppercase tracking-wider mb-1.5">
                      Número de inventario *
                    </label>
                    <select [(ngModel)]="mapeo.numero_inventario" class="input-gov text-sm">
                      <option value="">— Selecciona columna —</option>
                      @for (col of columnasExcel(); track col) {
                        <option [value]="col">{{ col }}</option>
                      }
                    </select>
                    @if (mapeo.numero_inventario && filasPrevia()[0]) {
                      <p class="text-xs text-guinda-400 mt-1 font-mono">
                        Ej: {{ filasPrevia()[0][mapeo.numero_inventario] }}
                      </p>
                    }
                  </div>

                  <div>
                    <label class="block text-xs font-semibold text-guinda-600 uppercase tracking-wider mb-1.5">
                      Descripción *
                    </label>
                    <select [(ngModel)]="mapeo.descripcion" class="input-gov text-sm">
                      <option value="">— Selecciona columna —</option>
                      @for (col of columnasExcel(); track col) {
                        <option [value]="col">{{ col }}</option>
                      }
                    </select>
                    @if (mapeo.descripcion && filasPrevia()[0]) {
                      <p class="text-xs text-guinda-400 mt-1 font-mono truncate">
                        Ej: {{ filasPrevia()[0][mapeo.descripcion] }}
                      </p>
                    }
                  </div>

                  <div>
                    <label class="block text-xs font-semibold text-guinda-600 uppercase tracking-wider mb-1.5">
                      Clasificación
                      <span class="text-guinda-300 normal-case font-normal">(opcional)</span>
                    </label>
                    <select [(ngModel)]="mapeo.clasificacion" class="input-gov text-sm">
                      <option value="">— No mapear —</option>
                      @for (col of columnasExcel(); track col) {
                        <option [value]="col">{{ col }}</option>
                      }
                    </select>
                  </div>

                  <div>
                    <label class="block text-xs font-semibold text-guinda-600 uppercase tracking-wider mb-1.5">
                      Ubicación esperada
                      <span class="text-guinda-300 normal-case font-normal">(opcional)</span>
                    </label>
                    <select [(ngModel)]="mapeo.ubicacion_esperada" class="input-gov text-sm">
                      <option value="">— No mapear —</option>
                      @for (col of columnasExcel(); track col) {
                        <option [value]="col">{{ col }}</option>
                      }
                    </select>
                  </div>
                </div>
              </div>

              <div class="flex gap-3">
                <button (click)="paso.set(1)" class="btn-ghost">Atrás</button>
                <button (click)="previsualizarImportacion()"
                        [disabled]="!mapeo.numero_inventario || !mapeo.descripcion"
                        class="btn-primary flex items-center gap-2 disabled:opacity-40 disabled:cursor-not-allowed">
                  Previsualizar importación
                  <svg class="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13.5 4.5L21 12m0 0l-7.5 7.5M21 12H3"/>
                  </svg>
                </button>
              </div>
            }

            @if (paso() === 3) {
              <div class="grid grid-cols-3 gap-4">
                <div class="stat-card border-t-2 border-t-success-DEFAULT text-center">
                  <span class="stat-label">Bienes válidos</span>
                  <span class="stat-value text-success-DEFAULT">{{ bienesValidos().length }}</span>
                </div>
                <div class="stat-card border-t-2 border-t-danger-DEFAULT text-center">
                  <span class="stat-label">Con errores</span>
                  <span class="stat-value text-danger-DEFAULT">{{ bienesConError().length }}</span>
                </div>
                <div class="stat-card border-t-2 border-t-guinda-300 text-center">
                  <span class="stat-label">Total filas</span>
                  <span class="stat-value text-guinda-600">{{ filasPrevia().length }}</span>
                </div>
              </div>

              @if (bienesConError().length > 0) {
                <div class="bg-danger-bg border border-danger-border rounded-lg p-4">
                  <p class="text-sm font-semibold text-danger-DEFAULT mb-2">
                    Filas omitidas por error:
                  </p>
                  <div class="space-y-1 max-h-28 overflow-y-auto">
                    @for (e of bienesConError().slice(0, 10); track $index) {
                      <p class="text-xs font-mono text-danger-DEFAULT">
                        Fila {{ e.fila }}: {{ e.error }}
                      </p>
                    }
                  </div>
                </div>
              }

              <div class="overflow-x-auto rounded-lg border border-crema-border">
                <table class="text-xs w-full">
                  <thead class="bg-guinda-50">
                    <tr>
                      <th class="px-3 py-2 text-left font-semibold text-guinda-600 border-b border-crema-border">No. Inventario</th>
                      <th class="px-3 py-2 text-left font-semibold text-guinda-600 border-b border-crema-border">Descripción</th>
                      <th class="px-3 py-2 text-left font-semibold text-guinda-600 border-b border-crema-border">Clasificación</th>
                      <th class="px-3 py-2 text-left font-semibold text-guinda-600 border-b border-crema-border">Ubicación</th>
                    </tr>
                  </thead>
                  <tbody>
                    @for (b of bienesValidos().slice(0, 6); track $index) {
                      <tr class="border-b border-crema-border last:border-0">
                        <td class="px-3 py-2 font-mono font-semibold text-guinda-700">{{ b.numero_inventario }}</td>
                        <td class="px-3 py-2 text-guinda-500 truncate max-w-48">{{ b.descripcion }}</td>
                        <td class="px-3 py-2 text-guinda-400">{{ b.clasificacion || '—' }}</td>
                        <td class="px-3 py-2 text-guinda-400 truncate max-w-36">{{ b.ubicacion_esperada || '—' }}</td>
                      </tr>
                    }
                  </tbody>
                </table>
                @if (bienesValidos().length > 6) {
                  <p class="text-xs text-guinda-400 text-center py-2 border-t border-crema-border">
                    ... y {{ bienesValidos().length - 6 }} bienes más
                  </p>
                }
              </div>

              @if (errorImport()) {
                <div class="flex items-center gap-2 bg-danger-bg border border-danger-border
                            rounded-lg px-4 py-3 animate-fade-in">
                  <span class="text-sm text-danger-DEFAULT">{{ errorImport() }}</span>
                </div>
              }

              @if (exitoImport()) {
                <div class="flex items-center gap-3 bg-success-bg border border-success-border
                            rounded-lg px-4 py-3 animate-fade-in">
                  <svg class="w-5 h-5 text-success-DEFAULT shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M4.5 12.75l6 6 9-13.5"/>
                  </svg>
                  <span class="text-sm text-success-DEFAULT font-semibold">{{ exitoImport() }}</span>
                </div>
              }

              <div class="flex gap-3">
                <button (click)="paso.set(2)" class="btn-ghost" [disabled]="importando()">Atrás</button>
                <button (click)="importar()"
                        [disabled]="bienesValidos().length === 0 || importando() || !!exitoImport()"
                        class="btn-primary flex-1 flex items-center justify-center gap-2
                               disabled:opacity-40 disabled:cursor-not-allowed">
                  @if (importando()) {
                    <svg class="w-4 h-4 animate-spin" fill="none" viewBox="0 0 24 24">
                      <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/>
                      <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/>
                    </svg>
                    Importando {{ bienesValidos().length }} bienes...
                  } @else if (exitoImport()) {
                    <svg class="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M4.5 12.75l6 6 9-13.5"/>
                    </svg>
                    Importado correctamente
                  } @else {
                    <svg class="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2"
                        d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5"/>
                    </svg>
                    Importar {{ bienesValidos().length }} bienes
                  }
                </button>
              </div>
            }
          </div>
        </div>
      </div>
    }
  `,
})
export class ImportadorCatalogoComponent {
  abierto = signal(false);
  paso = signal(1);
  archivo = signal<File | null>(null);
  columnasExcel = signal<string[]>([]);
  filasPrevia = signal<any[]>([]);
  errorArchivo = signal('');
  errorImport = signal('');
  exitoImport = signal('');
  importando = signal(false);
  bienesValidos = signal<any[]>([]);
  bienesConError = signal<any[]>([]);

  mapeo: MapeoColumnas = {
    numero_inventario: '',
    descripcion: '',
    clasificacion: '',
    ubicacion_esperada: '',
  };

  constructor(
    private api: ApiService,
    private cdr: ChangeDetectorRef,
  ) { }

  abrir() { this.abierto.set(true); this.reset(); }
  cerrar() { this.abierto.set(false); this.reset(); }

  reset() {
    this.paso.set(1);
    this.archivo.set(null);
    this.columnasExcel.set([]);
    this.filasPrevia.set([]);
    this.errorArchivo.set('');
    this.errorImport.set('');
    this.exitoImport.set('');
    this.bienesValidos.set([]);
    this.bienesConError.set([]);
    this.mapeo = { numero_inventario: '', descripcion: '', clasificacion: '', ubicacion_esperada: '' };
  }

  resetArchivo() {
    this.archivo.set(null);
    this.columnasExcel.set([]);
    this.filasPrevia.set([]);
    this.errorArchivo.set('');
  }

  onDrop(event: DragEvent) {
    event.preventDefault();
    const file = event.dataTransfer?.files[0];
    if (file) this.procesarArchivo(file);
  }

  onArchivoSeleccionado(event: Event) {
    const input = event.target as HTMLInputElement;
    const file = input.files?.[0];
    if (file) this.procesarArchivo(file);
  }

  procesarArchivo(file: File) {
    this.errorArchivo.set('');
    const ext = file.name.split('.').pop()?.toLowerCase();
    if (!['xlsx', 'xls', 'csv'].includes(ext ?? '')) {
      this.errorArchivo.set('Solo se aceptan archivos .xlsx, .xls o .csv');
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target!.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const filas = XLSX.utils.sheet_to_json(ws, { defval: '' }) as any[];

        if (filas.length === 0) {
          this.errorArchivo.set('El archivo no tiene datos');
          return;
        }

        const columnas = Object.keys(filas[0]);
        this.archivo.set(file);
        this.columnasExcel.set(columnas);
        this.filasPrevia.set(filas);
        this.autoDetectarColumnas(columnas);
        this.cdr.detectChanges();
      } catch {
        this.errorArchivo.set('No se pudo leer el archivo. Verifica que sea un Excel válido.');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  autoDetectarColumnas(cols: string[]) {
    const lower = cols.map(c => c.toLowerCase());
    const findCol = (keywords: string[]) =>
      cols.find((_, i) => keywords.some(k => lower[i].includes(k))) ?? '';

    this.mapeo.numero_inventario = findCol(['inventario', 'numero', 'código', 'codigo', 'no.', 'num', 'clave']);
    this.mapeo.descripcion = findCol(['descripcion', 'descripción', 'nombre', 'bien', 'articulo', 'artículo']);
    this.mapeo.clasificacion = findCol(['clasificacion', 'clasificación', 'tipo', 'categoria', 'categoría']);
    this.mapeo.ubicacion_esperada = findCol(['ubicacion', 'ubicación', 'lugar', 'area', 'área', 'localización', 'resguardo']);
  }

  previsualizarImportacion() {
    const validos: any[] = [];
    const errores: any[] = [];

    this.filasPrevia().forEach((fila, i) => {
      const num = String(fila[this.mapeo.numero_inventario] ?? '').trim();
      const desc = String(fila[this.mapeo.descripcion] ?? '').trim();

      if (!num) { errores.push({ fila: i + 2, error: 'Número de inventario vacío' }); return; }
      if (!desc) { errores.push({ fila: i + 2, error: `Descripción vacía (${num})` }); return; }

      validos.push({
        numero_inventario: num,
        descripcion: desc,
        clasificacion: this.mapeo.clasificacion ? String(fila[this.mapeo.clasificacion] ?? '').trim() : '',
        ubicacion_esperada: this.mapeo.ubicacion_esperada ? String(fila[this.mapeo.ubicacion_esperada] ?? '').trim() : '',
      });
    });

    this.bienesValidos.set(validos);
    this.bienesConError.set(errores);
    this.paso.set(3);
  }

  importar() {
    if (this.bienesValidos().length === 0 || this.importando()) return;
    this.errorImport.set('');
    this.importando.set(true);

    const hash = btoa(this.archivo()!.name + Date.now()).slice(0, 32);

    this.api.importarCatalogo({
      nombre_archivo: this.archivo()!.name,
      hash_archivo: hash,
      bienes: this.bienesValidos(),
    }).subscribe({
      next: res => {
        this.importando.set(false);
        if (res.ok && res.data) {
          this.exitoImport.set(
            `✓ ${res.data.total_bienes} bienes importados correctamente (versión ${res.data.numero_version})`
          );
        } else {
          this.errorImport.set(res.error?.message ?? 'Error al importar');
        }
        this.cdr.detectChanges();
      },
      error: () => {
        this.importando.set(false);
        this.errorImport.set('Error al conectar con el servidor');
        this.cdr.detectChanges();
      },
    });
  }
}