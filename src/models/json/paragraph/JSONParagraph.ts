import { Paragraph, Run, WIData, WIProperty, StyleOptions } from '../wordJsonModels';
import JSONRun from '../JSONRun';

export default class JSONParagraph {
  paragraphStyles: StyleOptions;
  paragraphTemplate: Paragraph;
  wiId: number;
  constructor(field: WIProperty, paragraphStyles: StyleOptions, wiId: number, headingLevel: number) {
    this.paragraphStyles = paragraphStyles;
    this.wiId = wiId;
    this.paragraphTemplate = {
      type: 'paragraph',
      runs: this.generateJsonParagraphRun(field, this.paragraphStyles),
    };
    if (headingLevel && field.name === 'Title:') {
      this.paragraphTemplate.headingLevel = headingLevel;
    } else {
      this.paragraphTemplate.headingLevel = 0;
    }
  } //constructor

  generateJsonParagraphRun(field: WIProperty, paragraphStyles): Run[] {
    let jsonRun: JSONRun;
    let runs: Run[] = [];
    let titleStyle = JSON.parse(JSON.stringify(paragraphStyles));

    titleStyle.isBold = false;
    titleStyle.IsUnderline = true;
    titleStyle.InsertSpace = true;
    titleStyle.InsertLineBreak = false;
    //adds ':' to field titles
    if (field.name !== '') {
      field.name += ':';
    }
    //exclude titles for fields that only show values
    if (
      field.name !== 'Title:' &&
      field.name !== 'ID:' &&
      field.name !== 'Description:' &&
      field.name !== 'text:' &&
      field.name !== 'pageBreak:'
    ) {
      paragraphStyles.Uri = field.url;
      jsonRun = new JSONRun(field.name, titleStyle);
      runs = [...runs, ...jsonRun.runs];

      // Spacer between label and value must NOT be underlined — otherwise the label's underline
      // visually bleeds into the gap before the value starts. 'RawText' bypasses striphtml, which
      // would otherwise trim this single space down to ''.
      const spacerStyle = { ...titleStyle, IsUnderline: false };
      jsonRun = new JSONRun(' ', spacerStyle, 'RawText');
      runs = [...runs, ...jsonRun.runs];

      paragraphStyles.InsertLineBreak = false;
    } //end if

    if (field.name === 'Title:') {
      paragraphStyles.Uri = field.url;
      paragraphStyles.InsertLineBreak = false;
      jsonRun = new JSONRun(field.value, paragraphStyles);
      runs = [...runs, ...jsonRun.runs];
    } else if (field.name === 'pageBreak:') {
      paragraphStyles.InsertPageBreak = true;
      jsonRun = new JSONRun(field.value, paragraphStyles);
      runs = [...runs, ...jsonRun.runs];
    } else {
      paragraphStyles.Uri = field.url;
      jsonRun = new JSONRun(field.value, paragraphStyles);
      runs = [...runs, ...jsonRun.runs];
    }
    return runs;
  } //generateJsonParagraph

  getJSONParagraph(): Paragraph {
    return this.paragraphTemplate;
  } //getJSONTable
} //class
