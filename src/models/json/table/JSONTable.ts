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

    data.forEach((rowData: WIData) => {
      let row = new JSONTableRow(rowData, tableStyles, retrieveOriginal, undefined, isFlattened);
      rows.push(row.getRow());
    });

    return rows;
  } //generateJsonRows

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
      level: data.level,
    };
  } //headersRowAdapter

  getJSONTable(): Table {
    return this.tableTemplate;
  } //getJSONTable
} //class
