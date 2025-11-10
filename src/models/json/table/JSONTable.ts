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
    }
  ) {
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

    // Apply vertical merging for grouped source cells (first row shows value; subsequent are empty)
    this.applyVerticalMerges(rows);

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
