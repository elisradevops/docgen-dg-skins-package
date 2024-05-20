import {
  TableRow,
  TableCell,
  WIData,
  WIProperty,
  StyleOptions
} from "../wordJsonModels";
import JSONTableCell from "./JSONTableCell";

export default class JSONTableRow {
  row: TableRow;

  constructor(data: WIData, tableStyles: StyleOptions) {
    this.row = this.generateJsonRow(data, tableStyles);
  } //constructor

  generateJsonRow(data: WIData, styles: StyleOptions): TableRow {
    let tableCells: TableCell[] = [];

    data.fields.forEach((wiProperty: WIProperty) => {
      if (wiProperty.name === "id") {
        styles.Uri = data.url;
      }
      let jsonTableCell = new JSONTableCell(wiProperty, styles);
      tableCells.push(JSON.parse(JSON.stringify(jsonTableCell.getJsonCell())));
    });

    return { Cells: tableCells };
  } //generateJsonRow

  getRow(): TableRow {
    return this.row;
  } //getRow
} //class
