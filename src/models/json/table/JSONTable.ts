import logger from '../../../services/logger';
import {
  Table,
  TableRow,
  WIQueryResults,
  WIData,
  WIProperty,
  StyleOptions,
  Shading,
} from '../wordJsonModels';
import JSONTableRow from './JSONTableRow';
export default class JSONTable {
  tableTemplate: Table;
  tableStyles: StyleOptions;
  private readonly enableVerticalMerge: boolean;
  private readonly enableAdaptiveLayout: boolean;

  constructor(
    data: WIQueryResults,
    headerRowStyle: StyleOptions,
    tableStyles: StyleOptions,
    headingLvl: number,
    retrieveOriginal = false,
    insertPageBreak: boolean = false,
    isFlattened = false,
    groupedHeader?: {
      leftLabel: string;
      rightLabel: string;
      leftColumns?: number;
      rightColumns?: number;
      shading?: Shading;
      leftShading?: Shading;
      rightShading?: Shading;
    },
    enableVerticalMerge: boolean = false,
    enableAdaptiveLayout: boolean = false
  ) {
    this.enableVerticalMerge = enableVerticalMerge;
    this.enableAdaptiveLayout = enableAdaptiveLayout;
    this.tableStyles = tableStyles;
    this.tableTemplate = {
      type: 'table',
      headingLevel: headingLvl,
      Rows: this.generateJsonRows(
        data,
        headerRowStyle,
        this.tableStyles,
        retrieveOriginal,
        isFlattened,
        groupedHeader
      ),
      insertPageBreak: insertPageBreak,
    };
  }

  generateJsonRows(
    data: any,
    headerRowStyle: StyleOptions,
    tableStyles: StyleOptions,
    retrieveOriginal: boolean,
    isFlattened: boolean,
    groupedHeader?: {
      leftLabel: string;
      rightLabel: string;
      leftColumns?: number;
      rightColumns?: number;
      shading?: Shading;
      leftShading?: Shading;
      rightShading?: Shading;
    }
  ): TableRow[] {
    let rows: TableRow[] = [];

    // A 0-result query is a normal, expected case (e.g. no previous open tasks right now) — the
    // column structure can't be derived without at least one row, so render a single informational
    // row instead of crashing (headersRowAdapter would otherwise throw reading data[0].fields).
    if (!Array.isArray(data) || data.length === 0) {
      return [this.buildNoResultsRow(tableStyles)];
    }

    // Scoped to callers that opt in (the generic query/table renderer) — leaves every other
    // table's hand-tuned layout (test-plan, SVD grouped headers, etc.) byte-identical.
    if (this.enableAdaptiveLayout && Array.isArray(data) && data.length > 0) {
      data = this.reorderPriorityColumns(data);
      data = this.dropEmptyColumns(data);
      const columnWidths = this.calculateColumnWidths(data);
      data.forEach((rowData: WIData) => {
        rowData.fields?.forEach((field: WIProperty, i: number) => {
          field.width = columnWidths[i];
        });
      });
    }

    let headersRowData: WIData = this.headersRowAdapter(data[0]);

    const defaultHeaderRowStyle = {
      isBold: true,
      IsItalic: false,
      IsUnderline: false,
      Size: 12,
      Uri: null,
      Font: 'Arial',
      InsertLineBreak: false,
      InsertSpace: false,
    };

    // Use the provided style or fall back to the default style
    const finalHeaderRowStyle = headerRowStyle ?? defaultHeaderRowStyle;

    // Determine if shading is needed
    const headerShading =
      headerRowStyle === undefined
        ? undefined
        : {
            color: 'auto',
            fill: '17365D',
            themeFillShade: 'BF',
          };

    // Optional grouped header row (two merged cells)
    const groupedRow = this.tryBuildGroupedHeaderRow(headersRowData, groupedHeader, finalHeaderRowStyle);
    if (groupedRow) {
      rows.push(groupedRow);
    }

    const headersRow = new JSONTableRow(headersRowData, finalHeaderRowStyle, undefined, headerShading);
    rows.push(headersRow.getRow());
    if (!data) {
      throw new Error('Missing table data');
    }
    data.forEach((rowData: WIData) => {
      let row = new JSONTableRow(rowData, tableStyles, retrieveOriginal, undefined, isFlattened);
      rows.push(row.getRow());
    });

    if (this.enableVerticalMerge) {
      // Apply vertical merging for grouped source cells (first row shows value; subsequent are empty)
      this.applyVerticalMerges(rows);
    }

    return rows;
  } //generateJsonRows

  // Builds a top grouped header row with two merged cells spanning left/right columns.
  // Returns null if no groupedHeader is provided.
  private tryBuildGroupedHeaderRow(
    headersRowData: WIData,
    groupedHeader: {
      leftLabel: string;
      rightLabel: string;
      leftColumns?: number;
      rightColumns?: number;
      shading?: Shading;
      leftShading?: Shading;
      rightShading?: Shading;
    },
    textStyle: StyleOptions
  ): TableRow | null {
    try {
      if (!groupedHeader || !groupedHeader.leftLabel || !groupedHeader.rightLabel) return null;

      const totalCols = headersRowData?.fields?.length || 0;
      if (totalCols === 0) return null;

      const leftCols = Math.max(1, groupedHeader.leftColumns ?? Math.floor(totalCols / 2));
      const rightCols = Math.max(1, groupedHeader.rightColumns ?? totalCols - leftCols);

      // Build simple table cells with Runs and gridSpan
      const makeCell = (label: string, span: number, cellShading?: Shading): any => {
        return {
          attachments: [],
          Paragraphs: [
            {
              Runs: [
                {
                  text: label,
                  Bold: !!textStyle?.isBold,
                  Italic: !!textStyle?.IsItalic,
                  Underline: !!textStyle?.IsUnderline,
                  Size: textStyle?.Size ?? 12,
                  Uri: null,
                  Font: textStyle?.Font ?? 'Arial',
                  InsertLineBreak: false,
                  InsertSpace: false,
                },
              ],
            },
          ],
          Html: undefined,
          width: '',
          shading: cellShading ?? groupedHeader?.shading,
          gridSpan: span,
        };
      };

      const leftShade = groupedHeader.leftShading ?? groupedHeader.shading;
      const rightShade = groupedHeader.rightShading ?? groupedHeader.shading;
      const leftCell = makeCell(groupedHeader.leftLabel, leftCols, leftShade);
      const rightCell = makeCell(groupedHeader.rightLabel, rightCols, rightShade);
      return { Cells: [leftCell, rightCell] } as TableRow;
    } catch {
      return null;
    }
  }

  private applyVerticalMerges(rows: TableRow[]) {
    if (!rows || rows.length <= 2) return; // at least headers + one data row

    const getCellText = (cell: any): string => {
      try {
        const htmlRaw = String(cell?.Html?.Html || '').trim();
        if (htmlRaw) {
          const htmlText = htmlRaw
            .replace(/<[^>]*>/g, ' ')
            .replace(/&nbsp;/gi, ' ')
            .replace(/\s+/g, ' ')
            .trim();
          if (htmlText) {
            return htmlText;
          }
        }
        const runs = cell?.Paragraphs?.[0]?.Runs || [];
        return runs.map((r: any) => (r?.text ?? '')).join('').trim();
      } catch {
        return '';
      }
    };

    // Determine how many header rows are present at the top (1 or 2 if grouped header exists)
    const hasGroupedHeader = Array.isArray(rows[0]?.Cells) && rows[0].Cells.some((c: any) => c?.gridSpan && c.gridSpan > 1);
    const dataRowStart = hasGroupedHeader ? 2 : 1; // skip grouped + column header
    const maxCols = Math.max(...rows.map((r) => (r?.Cells?.length || 0)));

    for (let col = 0; col < maxCols; col++) {
      for (let i = dataRowStart + 1; i < rows.length; i++) {
        const prev = rows[i - 1]?.Cells?.[col];
        const curr = rows[i]?.Cells?.[col];
        if (!prev || !curr) continue;

        const sameFill = (prev.shading?.fill || '') === (curr.shading?.fill || '');
        const prevTextEmpty = getCellText(prev) === '';
        const currTextEmpty = getCellText(curr) === '';

        if (!sameFill) {
          continue;
        }

        // Start merge when prev has text and current is empty
        if (!prevTextEmpty && currTextEmpty) {
          prev.vMerge = prev.vMerge ?? 'restart';
          curr.vMerge = 'continue';
        } else if ((prev.vMerge === 'restart' || prev.vMerge === 'continue') && currTextEmpty) {
          // Continue an existing merge block
          curr.vMerge = 'continue';
        }
      }
    }
  }

  // Single-cell placeholder row shown when the query returned zero work items — the column
  // structure can't be known without at least one row, so a plain message is the safest output.
  private buildNoResultsRow(tableStyles: StyleOptions): TableRow {
    return {
      Cells: [
        {
          attachments: [],
          Paragraphs: [
            {
              Runs: [
                {
                  text: 'No results found for this query.',
                  Bold: false,
                  Italic: true,
                  Underline: false,
                  Size: tableStyles?.Size ?? 12,
                  Uri: null as any,
                  Font: tableStyles?.Font ?? 'Arial',
                  InsertLineBreak: false,
                  InsertSpace: false,
                },
              ],
            },
          ],
          width: '',
        },
      ],
    };
  }

  // Common/important fields lead, matching how other doc-type adapters curate column order
  // (e.g. TraceQueryResultsSkinAdapter pushes Title first, CustomerCoverageTableSkinAdapter
  // hardcodes ID-then-Title) — the generic renderer otherwise just passes through whatever
  // order the ADO query definition happened to use.
  private static readonly PRIORITY_COLUMN_ORDER = ['id', 'title', 'work item type', 'state'];

  private reorderPriorityColumns(data: WIData[]): WIData[] {
    const fields = data[0]?.fields;
    if (!fields || fields.length === 0) return data;

    const priorityIndex = (name: string): number => {
      const normalized = String(name ?? '').trim().toLowerCase();
      return JSONTable.PRIORITY_COLUMN_ORDER.indexOf(normalized);
    };

    // Column order is uniform across rows, so compute the target index order once from the
    // first row and apply it identically to every row.
    const order = fields
      .map((field: WIProperty, i: number) => ({ i, priority: priorityIndex(field.name) }))
      .sort((a, b) => {
        const aHas = a.priority !== -1;
        const bHas = b.priority !== -1;
        if (aHas && bHas) return a.priority - b.priority;
        if (aHas) return -1;
        if (bHas) return 1;
        return a.i - b.i; // preserve original relative order for non-priority columns
      })
      .map((entry) => entry.i);

    if (order.every((idx, i) => idx === i)) return data; // already in priority order — no-op

    return data.map((wi) => ({
      ...wi,
      fields: order.map((idx) => wi.fields[idx]),
    }));
  }

  // Drops columns whose value is blank across every row (e.g. a Stack Rank field never populated
  // for the work item types in this query) — never drops every column.
  private dropEmptyColumns(data: WIData[]): WIData[] {
    const columnCount = data[0]?.fields?.length || 0;
    if (columnCount === 0) return data;

    const isColumnEmpty = new Array(columnCount).fill(true);
    data.forEach((wi) => {
      wi.fields?.forEach((field: WIProperty, i: number) => {
        if (i < columnCount && String(field?.value ?? '').trim() !== '') {
          isColumnEmpty[i] = false;
        }
      });
    });

    const emptyCount = isColumnEmpty.filter(Boolean).length;
    if (emptyCount === 0 || emptyCount === columnCount) return data;

    return data.map((wi) => ({
      ...wi,
      fields: wi.fields.filter((_: WIProperty, i: number) => !isColumnEmpty[i]),
    }));
  }

  // Content-driven column widths (percentage strings) — width per column scales with the longer
  // of its header label or its longest cell value, clamped and renormalized to sum to 100%.
  private calculateColumnWidths(data: WIData[]): string[] {
    const columnCount = data[0]?.fields?.length || 0;
    if (columnCount === 0) return [];

    const MIN_PCT = 8;
    const MAX_PCT = 35;

    const maxLenPerColumn = new Array(columnCount).fill(0);
    data.forEach((wi) => {
      wi.fields?.forEach((field: WIProperty, i: number) => {
        if (i >= columnCount) return;
        const labelLen = String(field?.name ?? '').length;
        const valueLen = String(field?.value ?? '').length;
        maxLenPerColumn[i] = Math.max(maxLenPerColumn[i], labelLen, valueLen);
      });
    });

    const totalLen = maxLenPerColumn.reduce((sum: number, len: number) => sum + len, 0) || 1;
    const clamped = maxLenPerColumn.map((len: number) =>
      Math.min(MAX_PCT, Math.max(MIN_PCT, (len / totalLen) * 100)),
    );

    const clampedTotal = clamped.reduce((sum: number, pct: number) => sum + pct, 0) || 1;
    return clamped.map((pct: number) => `${Number(((pct / clampedTotal) * 100).toFixed(1))}%`);
  }

  headersRowAdapter(data: WIData): WIData {
    let headerValuesWi;
    try {
      headerValuesWi = data.fields.map((field: WIProperty) => {
        return { name: 'header', value: field.name, width: field.width || '' };
      });
    } catch (error) {
      logger.error(`no fields to append`);
    }
    return {
      url: '',
      fields: headerValuesWi,
      Source: 999999999999,
      level: data?.level || 1,
    };
  } //headersRowAdapter

  getJSONTable(): Table {
    return this.tableTemplate;
  } //getJSONTable
} //class
