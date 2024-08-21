import logger from '../../../services/logger';
import { Table, TableRow, WIQueryResults, WIData, WIProperty, StyleOptions } from '../wordJsonModels';
import JSONTableRow from './JSONTableRow';
export default class JSONTable {
  tableTemplate: Table;
  tableStyles: StyleOptions;

  constructor(data: WIQueryResults, tableStyles: StyleOptions, headingLvl: number, retrieveOriginal = false) {
    this.tableStyles = tableStyles;
    this.tableTemplate = {
      type: 'table',
      headingLevel: headingLvl,
      Rows: this.generateJsonRows(data, this.tableStyles, retrieveOriginal),
    };
  }

  generateJsonRows(data: any, tableStyles: StyleOptions, retrieveOriginal: boolean): TableRow[] {
    let rows: TableRow[] = [];
    let headersRowData: WIData = this.headersRowAdapter(data[0]);

    let style2 = {
      isBold: true,
      IsItalic: false,
      IsUnderline: false,
      Size: 12,
      Uri: null,
      Font: 'Arial',
      InsertLineBreak: false,
      InsertSpace: false,
    };
    let headersRow = new JSONTableRow(headersRowData, style2);

    rows.push(headersRow.getRow());

    data.forEach((rowData: WIData) => {
      let row = new JSONTableRow(rowData, tableStyles, retrieveOriginal);
      rows.push(row.getRow());
    });

    return rows;
  } //generateJsonRows

  headersRowAdapter(data: WIData): WIData {
    let headerValuesWi;
    try {
      headerValuesWi = data.fields.map((field: WIProperty) => {
        return { name: 'header', value: field.name };
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
