"use strict";

import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import DataViewTable = powerbi.DataViewTable;
import DataViewTableRow = powerbi.DataViewTableRow;
import DataViewMetadataColumn = powerbi.DataViewMetadataColumn;

// ─── CSS ─────────────────────────────────────────────────────────────────────
const VISUAL_CSS = `
.fsem-wrap {
    display: flex;
    flex-direction: column;
    height: 100%;
    overflow: hidden;
    font-family: Calibri, 'Segoe UI', sans-serif;
    font-size: 12px;
    background: #fff;
}

.fsem-toolbar {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 6px 10px;
    background: #2E4153;
    flex-shrink: 0;
}

.fsem-title {
    color: #fff;
    font-weight: bold;
    font-size: 13px;
    flex: 1;
}

.fsem-btn {
    background: #3F5B70;
    color: #fff;
    border: none;
    padding: 5px 12px;
    border-radius: 4px;
    cursor: pointer;
    font-size: 11px;
    font-family: Calibri, sans-serif;
    white-space: nowrap;
}
.fsem-btn:hover { background: #4e6e85; }

.fsem-scroll {
    flex: 1;
    overflow: auto;
    min-height: 0;
}

.fsem-table {
    border-collapse: collapse;
    width: 100%;
    table-layout: auto;
}

/* ── Cabeçalho sticky ─────────────────────────────────────────────── */
.fsem-table thead {
    position: sticky;
    top: 0;
    z-index: 10;
}

.fsem-th-mes {
    background: #2E4153;
    color: #fff;
    font-weight: bold;
    text-align: center;
    padding: 7px 10px;
    font-size: 13px;
    border: 1px solid #2E4153;
}

.fsem-th-group {
    background: #374E65;
    color: #fff;
    font-weight: bold;
    text-align: center;
    padding: 5px 8px;
    border: 1px solid #4e6e85;
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

.fsem-th-semana {
    background: #3F5B70;
    color: #fff;
    text-align: center;
    padding: 5px 6px;
    border: 1px solid #4e6e85;
    min-width: 72px;
    white-space: nowrap;
}
.fsem-th-semana .sem-num {
    display: block;
    font-weight: bold;
    font-size: 13px;
    line-height: 1.2;
}
.fsem-th-semana .sem-range {
    display: block;
    font-weight: normal;
    font-size: 9px;
    opacity: 0.85;
    margin-top: 1px;
}

.fsem-th-info {
    background: #3F5B70;
    color: #fff;
    font-weight: bold;
    text-align: center;
    padding: 5px 8px;
    border: 1px solid #4e6e85;
    min-width: 90px;
    white-space: nowrap;
    font-size: 11px;
}

/* ── Corpo ────────────────────────────────────────────────────────── */
.fsem-table tbody tr:nth-child(odd)  td { background: #ffffff; }
.fsem-table tbody tr:nth-child(even) td { background: #f2f2f2; }
.fsem-table tbody tr:hover           td { background: #d1dce5 !important; }

.fsem-td-semana {
    border: 1px solid #ccc;
    text-align: center;
    padding: 5px 4px;
    font-weight: bold;
    font-size: 13px;
    color: #2E4153;
    white-space: nowrap;
}

.fsem-td-semana.compliance {
    background: #00B050 !important;
    color: #fff !important;
}

.fsem-td-semana.nao-compliance {
    background: #FF0000 !important;
    color: #fff !important;
}

.fsem-td-info {
    border: 1px solid #ccc;
    padding: 5px 8px;
    text-align: left;
    white-space: nowrap;
    color: #333;
    font-size: 12px;
}

/* ── Mensagem vazia ───────────────────────────────────────────────── */
.fsem-empty {
    padding: 40px;
    text-align: center;
    color: #888;
    font-size: 13px;
}
`;

// ─── helpers ─────────────────────────────────────────────────────────────────

function injectStyle(id: string, css: string): void {
    let el = document.getElementById(id);
    if (!el) {
        el = document.createElement("style");
        el.id = id;
        document.head.appendChild(el);
    }
    el.textContent = css;
}

/**
 * Extrai número da semana e intervalo de datas a partir do campo "Data Filtro".
 * Formatos suportados:
 *   "16 - 20/Abr a 26/Abr"
 *   "16 | 20/04 - 26/04"
 *   "16 – 20/04 a 26/04"
 */
function parseSemanaHeader(displayName: string): { num: string; range: string } {
    // Separador " | "
    let parts = displayName.split(" | ");
    if (parts.length === 2) return { num: parts[0].trim(), range: parts[1].trim() };

    // Separador " - " (primeira ocorrência)
    parts = displayName.split(" - ");
    if (parts.length >= 2) return { num: parts[0].trim(), range: parts.slice(1).join(" - ").trim() };

    // Separador " – " (em dash)
    parts = displayName.split(" – ");
    if (parts.length >= 2) return { num: parts[0].trim(), range: parts.slice(1).join(" – ").trim() };

    return { num: displayName, range: "" };
}

function escapeHtml(v: unknown): string {
    if (v === null || v === undefined) return "";
    return String(v)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

/**
 * Extrai o número inteiro de semana do rótulo, para ordenação.
 * Ex: "16 - 20/Abr a 26/Abr" → 16
 */
function semanaOrdinal(label: string): number {
    const { num } = parseSemanaHeader(label);
    const n = parseInt(num, 10);
    return isNaN(n) ? 9999 : n;
}

// ─── Tipos internos ────────────────────────────────────────────────────────────

interface PivotRow {
    /** Valores das colunas de info, indexados pelo index da coluna */
    info: Map<number, unknown>;
    /** Dias trabalhados por rótulo de semana */
    weekValues: Map<string, number | null>;
    /** Status compliance por rótulo de semana */
    complianceValues: Map<string, string>;
}

// ─── Visual principal ─────────────────────────────────────────────────────────

export class Visual implements IVisual {
    private host: powerbi.extensibility.visual.IVisualHost;
    private root: HTMLElement;
    private toolbar: HTMLElement;
    private scrollWrap: HTMLElement;

    // Dados para exportação
    private exportHeaders: string[] = [];
    private exportRows: string[][] = [];

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        injectStyle("fsem-style", VISUAL_CSS);

        this.root = document.createElement("div");
        this.root.className = "fsem-wrap";
        options.element.appendChild(this.root);

        // Toolbar
        this.toolbar = document.createElement("div");
        this.toolbar.className = "fsem-toolbar";

        const title = document.createElement("span");
        title.className = "fsem-title";
        title.textContent = "Frequência Semanal";
        this.toolbar.appendChild(title);

        const exportBtn = document.createElement("button");
        exportBtn.className = "fsem-btn";
        exportBtn.textContent = "📥 Exportar CSV";
        exportBtn.addEventListener("click", () => this.exportCSV());
        this.toolbar.appendChild(exportBtn);

        this.root.appendChild(this.toolbar);

        // Scroll container
        this.scrollWrap = document.createElement("div");
        this.scrollWrap.className = "fsem-scroll";
        this.root.appendChild(this.scrollWrap);
    }

    public update(options: VisualUpdateOptions): void {
        const dv = options?.dataViews?.[0];
        if (!dv?.table) {
            this.scrollWrap.innerHTML = `<div class="fsem-empty">Nenhum dado disponível. Adicione campos ao visual.</div>`;
            return;
        }
        this.render(dv.table);
    }

    private render(table: DataViewTable): void {
        const cols = table.columns;
        const rows = table.rows ?? [];

        // ── Classifica colunas por role ─────────────────────────────
        const weekKeyCols:    DataViewMetadataColumn[] = [];
        const weekValueCols:  DataViewMetadataColumn[] = [];
        const complianceCols: DataViewMetadataColumn[] = [];
        const infoCols:       DataViewMetadataColumn[] = [];

        cols.forEach(c => {
            if      (c.roles?.["weekKey"])    weekKeyCols.push(c);
            else if (c.roles?.["weekValue"])  weekValueCols.push(c);
            else if (c.roles?.["compliance"]) complianceCols.push(c);
            else                              infoCols.push(c);
        });

        // ── Índices das colunas ──────────────────────────────────────
        const weekKeyIdx   = weekKeyCols[0]?.index;    // coluna com rótulo da semana
        const weekValueIdx = weekValueCols[0]?.index;  // coluna com dias trabalhados
        const complianceIdx = complianceCols[0]?.index; // coluna com status (opcional)

        // Sem weekKey configurada, avisa o usuário
        if (weekKeyIdx === undefined || weekValueIdx === undefined) {
            this.scrollWrap.innerHTML = `<div class="fsem-empty">
                Configure os campos:<br>
                <strong>Chave da Semana</strong> (coluna "Data Filtro") e
                <strong>Valor da Semana</strong> (dias trabalhados).
            </div>`;
            return;
        }

        // ── Coleta semanas únicas e pivoteia ─────────────────────────
        // Chave de agrupamento por colaborador: concatena valores das infoCols
        const pivotMap = new Map<string, PivotRow>();
        const semanaLabels = new Set<string>();

        rows.forEach((row: DataViewTableRow) => {
            const semLabel = row[weekKeyIdx] != null ? String(row[weekKeyIdx]) : null;
            if (semLabel) semanaLabels.add(semLabel);

            // Chave única do colaborador = valores de info concatenados
            const infoKey = infoCols.map(ic => String(row[ic.index!] ?? "")).join("||");

            if (!pivotMap.has(infoKey)) {
                const infoMap = new Map<number, unknown>();
                infoCols.forEach(ic => infoMap.set(ic.index!, row[ic.index!]));
                pivotMap.set(infoKey, {
                    info: infoMap,
                    weekValues: new Map(),
                    complianceValues: new Map()
                });
            }

            const pivotRow = pivotMap.get(infoKey)!;

            if (semLabel) {
                const val = row[weekValueIdx];
                const numVal = val !== null && val !== undefined ? Number(val) : null;
                pivotRow.weekValues.set(semLabel, numVal);

                if (complianceIdx !== undefined) {
                    const status = String(row[complianceIdx] ?? "");
                    pivotRow.complianceValues.set(semLabel, status);
                }
            }
        });

        // Ordena semanas pelo número extraído
        const sortedSemanas = Array.from(semanaLabels).sort(
            (a, b) => semanaOrdinal(a) - semanaOrdinal(b)
        );

        const pivotRows = Array.from(pivotMap.values());
        const totalWeekCols = sortedSemanas.length;
        const totalInfoCols = infoCols.length;
        const totalCols     = totalWeekCols + totalInfoCols;

        // ── Extrai mês/ano do rótulo da primeira semana ──────────────
        let mesLabel = "";
        if (sortedSemanas.length > 0) {
            const header = parseSemanaHeader(sortedSemanas[0]);
            // Tenta padrão "dd/MMM" (ex: 20/Abr) ou "dd/MM" (ex: 20/04)
            const matchMes = header.range.match(/\d+\/([A-Za-zÀ-ú]+|\d+)/);
            if (matchMes) {
                const mesStr = matchMes[1];
                // Se for número, converte para nome
                const mesesNome = ["","Janeiro","Fevereiro","Março","Abril","Maio","Junho",
                                   "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
                const mesesAbrev = ["jan","fev","mar","abr","mai","jun",
                                    "jul","ago","set","out","nov","dez"];
                const numMes = parseInt(mesStr, 10);
                if (!isNaN(numMes)) {
                    mesLabel = `${mesesNome[numMes] || mesStr} / ${new Date().getFullYear()}`;
                } else {
                    const idx = mesesAbrev.indexOf(mesStr.toLowerCase().substring(0, 3));
                    mesLabel = idx >= 0
                        ? `${mesesNome[idx + 1]} / ${new Date().getFullYear()}`
                        : mesStr;
                }
            }
        }

        // ── Build HTML ────────────────────────────────────────────────
        const parts: string[] = [];
        parts.push(`<table class="fsem-table"><thead>`);

        // Linha 1: mês
        if (mesLabel) {
            parts.push(`<tr><th class="fsem-th-mes" colspan="${totalCols}">${escapeHtml(mesLabel)}</th></tr>`);
        }

        // Linha 2: grupos
        if (totalCols > 0) {
            parts.push(`<tr>`);
            if (totalWeekCols > 0)
                parts.push(`<th class="fsem-th-group" colspan="${totalWeekCols}">Semanas</th>`);
            if (totalInfoCols > 0)
                parts.push(`<th class="fsem-th-group" colspan="${totalInfoCols}">Dados do Colaborador</th>`);
            parts.push(`</tr>`);
        }

        // Linha 3: cabeçalhos individuais
        parts.push(`<tr>`);
        sortedSemanas.forEach(label => {
            const h = parseSemanaHeader(label);
            parts.push(
                `<th class="fsem-th-semana">` +
                `<span class="sem-num">${escapeHtml(h.num)}</span>` +
                (h.range ? `<span class="sem-range">${escapeHtml(h.range)}</span>` : "") +
                `</th>`
            );
        });
        infoCols.forEach(c => {
            parts.push(`<th class="fsem-th-info">${escapeHtml(c.displayName)}</th>`);
        });
        parts.push(`</tr></thead><tbody>`);

        // ── Corpo ─────────────────────────────────────────────────────
        if (pivotRows.length === 0) {
            parts.push(`<tr><td colspan="${totalCols}" class="fsem-empty">Sem registros para o período selecionado.</td></tr>`);
        } else {
            pivotRows.forEach(pr => {
                parts.push(`<tr>`);

                sortedSemanas.forEach(label => {
                    const val = pr.weekValues.get(label);
                    const hasData = val !== null && val !== undefined && !isNaN(val as number) && (val as number) > 0;

                    let cssClass = "";
                    if (hasData) {
                        const status = (pr.complianceValues.get(label) ?? "").toLowerCase();
                        if (status.includes("não compliance") || status.includes("nao compliance")) {
                            cssClass = "nao-compliance";
                        } else if (status.includes("compliance") || status === "") {
                            // verde se tem dados e não tem status negativo
                            cssClass = "compliance";
                        }
                    }

                    const display = hasData ? escapeHtml(val) : "";
                    parts.push(`<td class="fsem-td-semana ${cssClass}">${display}</td>`);
                });

                infoCols.forEach(ic => {
                    parts.push(`<td class="fsem-td-info">${escapeHtml(pr.info.get(ic.index!))}</td>`);
                });

                parts.push(`</tr>`);
            });
        }

        parts.push(`</tbody></table>`);
        this.scrollWrap.innerHTML = parts.join("");

        // ── Prepara dados para CSV ────────────────────────────────────
        this.exportHeaders = [
            ...sortedSemanas,
            ...infoCols.map(c => c.displayName)
        ];
        this.exportRows = pivotRows.map(pr => [
            ...sortedSemanas.map(label => {
                const val = pr.weekValues.get(label);
                return val !== null && val !== undefined ? String(val) : "";
            }),
            ...infoCols.map(ic => {
                const v = pr.info.get(ic.index!);
                return v !== null && v !== undefined ? String(v) : "";
            })
        ]);
    }

    // ── Exportação CSV ────────────────────────────────────────────────
    private exportCSV(): void {
        if (this.exportHeaders.length === 0) return;

        const escape = (v: string) => `"${v.replace(/"/g, '""')}"`;
        const lines: string[] = [];

        lines.push(this.exportHeaders.map(escape).join(";"));
        this.exportRows.forEach(row => lines.push(row.map(escape).join(";")));

        const bom  = "\uFEFF";
        const blob = new Blob([bom + lines.join("\n")], { type: "text/csv;charset=utf-8;" });
        const url  = URL.createObjectURL(blob);
        const a    = document.createElement("a");
        a.href     = url;
        a.download = `Frequencia_Semanal_${Date.now()}.csv`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }
}