import { TableRow, TableCell, WIData, WIProperty, StyleOptions } from '../wordJsonModels';
import JSONTableCell from './JSONTableCell';

export default class JSONTableRow {
  row: TableRow;

  constructor(data: WIData, tableStyles: StyleOptions, retrieveOriginal = false) {
    this.row = this.generateJsonRow(data, tableStyles, retrieveOriginal);
  } //constructor

  generateJsonRow(data: WIData, styles: StyleOptions, retrieveOriginal: boolean): TableRow {
    let tableCells: TableCell[] = [];

    data.fields.forEach((wiProperty: WIProperty) => {
      if (wiProperty.name === 'id') {
        styles.Uri = data.url;
      }
      let jsonTableCell = new JSONTableCell(wiProperty, styles, retrieveOriginal);
      tableCells.push(JSON.parse(JSON.stringify(jsonTableCell.getJsonCell())));
    });

    return { Cells: tableCells };
  } //generateJsonRow

  getRow(): TableRow {
    return this.row;
  } //getRow
} //class
