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

.fsem-td-semana.clickable {
    cursor: pointer;
}

.fsem-td-semana.clickable:hover {
    filter: brightness(0.88);
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

.fsem-empty {
    padding: 40px;
    text-align: center;
    color: #888;
    font-size: 13px;
}

/* ── Modal ────────────────────────────────────────────────────────── */
.fsem-overlay {
    display: none;
    position: fixed;
    top: 0; left: 0;
    width: 100%; height: 100%;
    background: rgba(0,0,0,0.45);
    z-index: 99998;
}

.fsem-modal {
    display: none;
    position: fixed;
    top: 50%; left: 50%;
    transform: translate(-50%, -50%);
    width: 720px;
    max-width: 95vw;
    max-height: 80vh;
    overflow-y: auto;
    background: #2E4153;
    border: 1px solid #3F5B70;
    border-radius: 10px;
    padding: 14px;
    z-index: 99999;
    box-shadow: 0 8px 32px rgba(0,0,0,0.5);
    font-family: Calibri, sans-serif;
    font-size: 12px;
    color: #fff;
}

.fsem-modal-header {
    display: flex;
    align-items: center;
    margin-bottom: 10px;
}

.fsem-modal-title {
    flex: 1;
    font-weight: bold;
    font-size: 13px;
    color: #fff;
}

.fsem-modal-close {
    background: #3F5B70;
    border: none;
    color: #fff;
    border-radius: 50%;
    width: 26px;
    height: 26px;
    font-size: 14px;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
}
.fsem-modal-close:hover { background: #4e6e85; }

.fsem-modal-table {
    border-collapse: collapse;
    width: 100%;
}

.fsem-modal-table th {
    background: #3F5B70;
    color: #fff;
    padding: 5px 8px;
    border: 1px solid #4e6e85;
    font-weight: bold;
    white-space: nowrap;
    position: sticky;
    top: 0;
}

.fsem-modal-table td {
    padding: 4px 7px;
    border: 1px solid #3F5B70;
    background: #2E4153;
    color: #fff;
    white-space: nowrap;
}

.fsem-modal-table tr:nth-child(even) td { background: #364d60; }
.fsem-modal-table tr:hover td { background: #3F5B70; }
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

function parseSemanaHeader(displayName: string): { num: string; range: string } {
    let parts = displayName.split(" | ");
    if (parts.length === 2) return { num: parts[0].trim(), range: parts[1].trim() };

    parts = displayName.split(" - ");
    if (parts.length >= 2) return { num: parts[0].trim(), range: parts.slice(1).join(" - ").trim() };

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

function semanaOrdinal(label: string): number {
    const { num } = parseSemanaHeader(label);
    const n = parseInt(num, 10);
    return isNaN(n) ? 9999 : n;
}

// ─── Tipos internos ───────────────────────────────────────────────────────────

interface DetailRow {
    values: Map<number, unknown>;
}

interface PivotRow {
    info: Map<number, unknown>;
    weekValues: Map<string, number | null>;
    complianceValues: Map<string, string>;
    details: Map<string, DetailRow[]>; // semanaLabel → linhas de detalhe
}

// ─── Visual principal ─────────────────────────────────────────────────────────

export class Visual implements IVisual {
    private host: powerbi.extensibility.visual.IVisualHost;
    private root: HTMLElement;
    private toolbar: HTMLElement;
    private scrollWrap: HTMLElement;
    private overlay: HTMLElement;
    private modal: HTMLElement;
    private modalTitle: HTMLElement;
    private modalBody: HTMLElement;

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

        // Botão Exportar CSV (bloqueado por permissão de admin — mantido comentado)
        const exportBtn = document.createElement("button");
        exportBtn.className = "fsem-btn";
        exportBtn.textContent = "📥 Exportar CSV";
        exportBtn.addEventListener("click", () => this.exportCSV());
        this.toolbar.appendChild(exportBtn);

        // Botão Copiar CSV
        const localBtn = document.createElement("button");
        localBtn.className = "fsem-btn";
        localBtn.textContent = "📋 Copiar CSV";
        localBtn.addEventListener("click", () => this.exportLocal());
        this.toolbar.appendChild(localBtn);

        this.root.appendChild(this.toolbar);

        // Scroll container
        this.scrollWrap = document.createElement("div");
        this.scrollWrap.className = "fsem-scroll";
        this.root.appendChild(this.scrollWrap);

        // Modal overlay
        this.overlay = document.createElement("div");
        this.overlay.className = "fsem-overlay";
        this.overlay.addEventListener("click", () => this.closeModal());
        this.root.appendChild(this.overlay);

        // Modal box
        this.modal = document.createElement("div");
        this.modal.className = "fsem-modal";

        const modalHeader = document.createElement("div");
        modalHeader.className = "fsem-modal-header";

        this.modalTitle = document.createElement("span");
        this.modalTitle.className = "fsem-modal-title";
        modalHeader.appendChild(this.modalTitle);

        const closeBtn = document.createElement("button");
        closeBtn.className = "fsem-modal-close";
        closeBtn.textContent = "✕";
        closeBtn.addEventListener("click", () => this.closeModal());
        modalHeader.appendChild(closeBtn);

        this.modal.appendChild(modalHeader);

        this.modalBody = document.createElement("div");
        this.modal.appendChild(this.modalBody);

        this.root.appendChild(this.modal);

        // Fecha com ESC
        document.addEventListener("keydown", (e) => {
            if (e.key === "Escape") this.closeModal();
        });
    }

    private openModal(title: string, detailRows: DetailRow[], detailCols: DataViewMetadataColumn[]): void {
        this.modalTitle.textContent = title;

        if (detailRows.length === 0) {
            this.modalBody.innerHTML = `<p style="color:#aaa;text-align:center;padding:20px;">Sem registros de detalhe para este período.</p>`;
        } else {
            const parts: string[] = [];
            parts.push(`<table class="fsem-modal-table"><thead><tr>`);
            detailCols.forEach(c => {
                parts.push(`<th>${escapeHtml(c.displayName)}</th>`);
            });
            parts.push(`</tr></thead><tbody>`);

            detailRows.forEach(dr => {
                parts.push(`<tr>`);
                detailCols.forEach(c => {
                    parts.push(`<td>${escapeHtml(dr.values.get(c.index!))}</td>`);
                });
                parts.push(`</tr>`);
            });

            parts.push(`</tbody></table>`);
            this.modalBody.innerHTML = parts.join("");
        }

        this.overlay.style.display = "block";
        this.modal.style.display = "block";
    }

    private closeModal(): void {
        this.overlay.style.display = "none";
        this.modal.style.display = "none";
    }

    public update(options: VisualUpdateOptions): void {
        const dv = options?.dataViews?.[0];
        if (!dv?.table) {
            this.scrollWrap.innerHTML = `<div class="fsem-empty">Nenhum dado disponível. Adicione campos ao visual.</div>`;
            return;
        }

        const table = dv.table;

        const firstInfoIndex = table.columns.findIndex(
            c => c.roles && c.roles["info"]
        );

        const sortedTable: powerbi.DataViewTable = {
            columns: table.columns,
            rows: firstInfoIndex >= 0
                ? [...(table.rows ?? [])].sort((a, b) => {
                    const valA = String(a[firstInfoIndex] ?? "").toLowerCase();
                    const valB = String(b[firstInfoIndex] ?? "").toLowerCase();
                    return valA.localeCompare(valB, "pt-BR");
                })
                : [...(table.rows ?? [])],
            totals: table.totals
        };

        this.render(sortedTable);
    }

    private render(table: DataViewTable): void {
        const cols = table.columns;
        const rows = table.rows ?? [];

        // ── Classifica colunas por role ──────────────────────────────
        const weekKeyCols: DataViewMetadataColumn[] = [];
        const weekValueCols: DataViewMetadataColumn[] = [];
        const complianceCols: DataViewMetadataColumn[] = [];
        const infoCols: DataViewMetadataColumn[] = [];
        const detailCols: DataViewMetadataColumn[] = [];

        cols.forEach(c => {
            if (c.roles?.["weekKey"]) weekKeyCols.push(c);
            else if (c.roles?.["weekValue"]) weekValueCols.push(c);
            else if (c.roles?.["compliance"]) complianceCols.push(c);
            else if (c.roles?.["detail"]) detailCols.push(c);
            else if (c.roles?.["info"]) infoCols.push(c);
        });

        const weekKeyIdx = weekKeyCols[0]?.index;
        const weekValueIdx = weekValueCols[0]?.index;
        const complianceIdx = complianceCols[0]?.index;

        if (weekKeyIdx === undefined || weekValueIdx === undefined) {
            this.scrollWrap.innerHTML = `<div class="fsem-empty">
                Configure os campos:<br>
                <strong>Chave da Semana</strong> e <strong>Valor da Semana</strong>.
            </div>`;
            return;
        }

        // ── Pivoteia dados ───────────────────────────────────────────
        const pivotMap = new Map<string, PivotRow>();
        const semanaLabels = new Set<string>();

        rows.forEach((row: DataViewTableRow) => {
            const semLabel = row[weekKeyIdx] != null ? String(row[weekKeyIdx]) : null;
            if (semLabel) semanaLabels.add(semLabel);

            const infoKey = infoCols.map(ic => String(row[ic.index!] ?? "")).join("||");

            if (!pivotMap.has(infoKey)) {
                const infoMap = new Map<number, unknown>();
                infoCols.forEach(ic => infoMap.set(ic.index!, row[ic.index!]));
                pivotMap.set(infoKey, {
                    info: infoMap,
                    weekValues: new Map(),
                    complianceValues: new Map(),
                    details: new Map()
                });
            }

            const pivotRow = pivotMap.get(infoKey)!;

            if (semLabel) {
                const val = row[weekValueIdx];
                const numVal = val !== null && val !== undefined ? Number(val) : null;
                pivotRow.weekValues.set(semLabel, numVal);

                if (complianceIdx !== undefined) {
                    pivotRow.complianceValues.set(semLabel, String(row[complianceIdx] ?? ""));
                }

                // Coleta detalhe se houver colunas de detail
                if (detailCols.length > 0) {
                    const detailMap = new Map<number, unknown>();
                    detailCols.forEach(dc => detailMap.set(dc.index!, row[dc.index!]));
                    const existing = pivotRow.details.get(semLabel) ?? [];
                    existing.push({ values: detailMap });
                    pivotRow.details.set(semLabel, existing);
                }
            }
        });

        const sortedSemanas = Array.from(semanaLabels).sort(
            (a, b) => semanaOrdinal(a) - semanaOrdinal(b)
        );

        const pivotRows = Array.from(pivotMap.values());
        const totalWeekCols = sortedSemanas.length;
        const totalInfoCols = infoCols.length;
        const totalCols = totalWeekCols + totalInfoCols;

        // ── Extrai mês ───────────────────────────────────────────────
        let mesLabel = "";
        if (sortedSemanas.length > 0) {
            const header = parseSemanaHeader(sortedSemanas[0]);
            const matchMes = header.range.match(/\d+\/([A-Za-zÀ-ú]+|\d+)/);
            if (matchMes) {
                const mesStr = matchMes[1];
                const mesesNome = ["", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
                const mesesAbrev = ["jan", "fev", "mar", "abr", "mai", "jun",
                    "jul", "ago", "set", "out", "nov", "dez"];
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

        // ── Build HTML ───────────────────────────────────────────────
        const parts: string[] = [];
        parts.push(`<table class="fsem-table"><thead>`);

        if (mesLabel) {
            parts.push(`<tr><th class="fsem-th-mes" colspan="${totalCols}">${escapeHtml(mesLabel)}</th></tr>`);
        }

        if (totalCols > 0) {
            parts.push(`<tr>`);
            if (totalWeekCols > 0)
                parts.push(`<th class="fsem-th-group" colspan="${totalWeekCols}">Semanas</th>`);
            if (totalInfoCols > 0)
                parts.push(`<th class="fsem-th-group" colspan="${totalInfoCols}">Dados do Colaborador</th>`);
            parts.push(`</tr>`);
        }

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

        // ── Corpo ────────────────────────────────────────────────────
        if (pivotRows.length === 0) {
            parts.push(`<tr><td colspan="${totalCols}" class="fsem-empty">Sem registros para o período selecionado.</td></tr>`);
        } else {
            pivotRows.forEach((pr, rowIdx) => {
                parts.push(`<tr>`);

                sortedSemanas.forEach((label, colIdx) => {
                    const val = pr.weekValues.get(label);
                    const hasData = val !== null && val !== undefined && !isNaN(val as number) && (val as number) > 0;
                    const hasDetail = detailCols.length > 0 && hasData;

                    let cssClass = "";
                    if (hasData) {
                        const status = (pr.complianceValues.get(label) ?? "").toLowerCase();
                        if (status.includes("não compliance") || status.includes("nao compliance")) {
                            cssClass = "nao-compliance";
                        } else if (status.includes("compliance") || status === "") {
                            cssClass = "compliance";
                        }
                    }

                    const clickable = hasDetail ? " clickable" : "";
                    const dataAttr = hasDetail
                        ? ` data-row="${rowIdx}" data-col="${colIdx}"`
                        : "";

                    const display = hasData ? escapeHtml(val) : "";
                    parts.push(`<td class="fsem-td-semana${cssClass ? " " + cssClass : ""}${clickable}"${dataAttr}>${display}</td>`);
                });

                infoCols.forEach(ic => {
                    parts.push(`<td class="fsem-td-info">${escapeHtml(pr.info.get(ic.index!))}</td>`);
                });

                parts.push(`</tr>`);
            });
        }

        parts.push(`</tbody></table>`);
        this.scrollWrap.innerHTML = parts.join("");

        // ── Event listener para popup ────────────────────────────────
        if (detailCols.length > 0) {
            this.scrollWrap.addEventListener("click", (e) => {
                const td = (e.target as HTMLElement).closest("td[data-row]") as HTMLElement;
                if (!td) return;

                const rowIdx = parseInt(td.getAttribute("data-row") ?? "-1", 10);
                const colIdx = parseInt(td.getAttribute("data-col") ?? "-1", 10);

                if (rowIdx < 0 || colIdx < 0) return;

                const pr = pivotRows[rowIdx];
                const semLabel = sortedSemanas[colIdx];
                const details = pr.details.get(semLabel) ?? [];

                const nomePessoa = infoCols.length > 0
                    ? String(pr.info.get(infoCols[0].index!) ?? "")
                    : "";
                const { num } = parseSemanaHeader(semLabel);

                this.openModal(
                    `${nomePessoa} — Semana ${num}`,
                    details,
                    detailCols
                );
            });
        }

        // ── Prepara dados para CSV ───────────────────────────────────
        const sortedPivotRows = [...pivotRows].sort((a, b) => {
            const firstInfoCol = infoCols[0];
            if (!firstInfoCol) return 0;
            const valA = String(a.info.get(firstInfoCol.index!) ?? "").toLowerCase();
            const valB = String(b.info.get(firstInfoCol.index!) ?? "").toLowerCase();
            return valA.localeCompare(valB, "pt-BR");
        });

        this.exportHeaders = [
            ...sortedSemanas.map(label => {
                const h = parseSemanaHeader(label);
                return h.range ? `${h.num} - ${h.range}` : h.num;
            }),
            ...infoCols.map(c => c.displayName)
        ];

        this.exportRows = sortedPivotRows.map(pr => [
            ...sortedSemanas.map(label => {
                const val = pr.weekValues.get(label);
                return (val !== null && val !== undefined && !isNaN(val as number) && (val as number) > 0)
                    ? String(val)
                    : "";
            }),
            ...infoCols.map(ic => {
                const v = pr.info.get(ic.index!);
                return v !== null && v !== undefined ? String(v) : "";
            })
        ]);
    }

    // ── Exportar CSV (bloqueado por permissão de admin) ───────────────
    private exportCSV(): void {
        // if (this.exportHeaders.length === 0) return;
        // ... comentado até permissão ser liberada
    }

    // ── Copiar CSV para clipboard ─────────────────────────────────────
    private exportLocal(): void {
        if (this.exportHeaders.length === 0) return;

        const escape = (v: string) => `"${v.replace(/"/g, '""')}"`;
        const lines: string[] = [];

        lines.push(this.exportHeaders.map(escape).join(";"));
        this.exportRows.forEach(row => lines.push(row.map(escape).join(";")));

        const content = lines.join("\n");

        navigator.clipboard.writeText(content).then(() => {
            const btn = this.toolbar.querySelector(".fsem-btn:last-child") as HTMLButtonElement;
            if (btn) {
                const original = btn.textContent;
                btn.textContent = "✅ Copiado!";
                btn.style.background = "#00B050";
                setTimeout(() => {
                    btn.textContent = original;
                    btn.style.background = "";
                }, 2000);
            }
        }).catch(() => {
            const ta = document.createElement("textarea");
            ta.value = content;
            ta.style.position = "fixed";
            ta.style.opacity = "0";
            document.body.appendChild(ta);
            ta.focus();
            ta.select();
            document.execCommand("copy");
            document.body.removeChild(ta);
        });
    }
}