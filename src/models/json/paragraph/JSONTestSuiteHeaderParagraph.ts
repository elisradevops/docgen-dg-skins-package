import {
  Paragraph,
  Run,
  WIData,
  WIProperty,
  StyleOptions,
} from "../wordJsonModels";
import JSONRun from "../JSONRun";

export default class JSONTestSuiteHeaderParagraph {
  paragraphStyles: StyleOptions;
  paragraphTemplate: Paragraph;
  wiId: number;

  constructor(
    fields: WIProperty[],
    paragraphStyles: StyleOptions,
    wiId: number,
    headingLevel: number
  ) {
    this.paragraphStyles = paragraphStyles;
    this.wiId = wiId;
    this.paragraphTemplate = {
      type: "paragraph",
      runs: this.generateJsonParagraphHeaderRuns(fields, this.paragraphStyles),
    };

    if (headingLevel) {
      this.paragraphTemplate.headingLevel = headingLevel;
    } else {
      this.paragraphTemplate.headingLevel = 0;
    }
  } //constructor

  generateJsonParagraphHeaderRuns(fields, paragraphStyles) {
    let runs: any = [];
    fields.forEach((field) => {
      runs = [
        ...runs,
        ...this.generateJsonParagraphRun(field, this.paragraphStyles),
      ];
    });

    return runs;
  }
  generateJsonParagraphRun(field: WIProperty, paragraphStyles): Run[] {
    let jsonRun: JSONRun;
    let runs: Run[] = [];
    let titleStyle = JSON.parse(JSON.stringify(paragraphStyles));

    titleStyle.isBold = false;
    titleStyle.IsUnderline = true;
    titleStyle.InsertSpace = true;
    titleStyle.InsertLineBreak = false;
    //adds ':' to field titles
    if (field.name !== "") {
      field.name += ": ";
    }
    //exclude titles for fields that only show values
    if (
      field.name !== "Title: " &&
      field.name !== "ID: " &&
      field.name !== "Description: "
    ) {
      paragraphStyles.Uri = field.url;
      jsonRun = new JSONRun(field.name, titleStyle);
      runs = [...runs, ...jsonRun.runs];

      paragraphStyles.InsertLineBreak = false;
    } //end if

    if (field.name === "ID: ") {
      field.value = `${field.value}`;
      paragraphStyles.Uri = field.url;
      paragraphStyles.InsertLineBreak = false;
    }

    if (field.name === "Title: ") {
      let fieldtype = "SuiteHeaderParagraphTitle";
      paragraphStyles.InsertLineBreak = false;
      paragraphStyles.Uri = field.url;
      jsonRun = new JSONRun(field.value, paragraphStyles, fieldtype);
      runs = [...runs, ...jsonRun.runs];
    } else {
      jsonRun = new JSONRun(field.value, paragraphStyles);
      runs = [...runs, ...jsonRun.runs];
    }
    return runs;
  } //generateJsonParagraph

  getJSONParagraph(): Paragraph {
    return this.paragraphTemplate;
  } //getJSONTable
} //class
