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
    isFlattened = false
  ) {
    this.tableStyles = tableStyles;
    this.tableTemplate = {
      type: 'table',
      headingLevel: headingLvl,
      Rows: this.generateJsonRows(data, headerRowStyle, this.tableStyles, retrieveOriginal, isFlattened),
      insertPageBreak: insertPageBreak,
    };
  }

  generateJsonRows(
    data: any,
    headerRowStyle: StyleOptions,
    tableStyles: StyleOptions,
    retrieveOriginal: boolean,
    isFlattened: boolean
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

  private applyVerticalMerges(rows: TableRow[]) {
    if (!rows || rows.length <= 2) return; // header + at least 1 data row needed

    const getCellText = (cell: any): string => {
      try {
        const runs = cell?.Paragraphs?.[0]?.Runs || [];
        return runs.map((r: any) => (r?.text ?? '')).join('').trim();
      } catch {
        return '';
      }
    };

    const dataRowStart = 1; // skip header
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
