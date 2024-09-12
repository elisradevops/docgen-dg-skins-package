import { TableRow, TableCell, WIData, WIProperty, StyleOptions, Shading } from '../wordJsonModels';
import JSONTableCell from './JSONTableCell';

export default class JSONTableRow {
  row: TableRow;

  constructor(
    data: WIData,
    tableStyles: StyleOptions,
    retrieveOriginal = false,
    shading: Shading = undefined
  ) {
    this.row = this.generateJsonRow(data, tableStyles, retrieveOriginal, shading);
  } //constructor

  generateJsonRow(data: WIData, styles: StyleOptions, retrieveOriginal: boolean, shading: Shading): TableRow {
    let tableCells: TableCell[] = [];

    data.fields.forEach((wiProperty: WIProperty) => {
      if (wiProperty.name === 'id') {
        styles.Uri = data.url;
      }
      let jsonTableCell = new JSONTableCell(wiProperty, styles, retrieveOriginal, shading);
      tableCells.push(JSON.parse(JSON.stringify(jsonTableCell.getJsonCell())));
    });

    return { Cells: tableCells };
  } //generateJsonRow

  getRow(): TableRow {
    return this.row;
  } //getRow
} //class
