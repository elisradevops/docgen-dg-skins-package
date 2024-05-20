import { Run, WIProperty, StyleOptions } from "../wordJsonModels";
import JSONHtml from "../html/JSONHtml";
import JSONParagraph from "./JSONParagraph";
import JSONTable from "../table/JSONTable";
import logger from "../../../services/logger";

export default class JSONRichTextParagraph {
  paragraphStyles: StyleOptions;
  paragraphTemplate: any;
  wiId: number;
  constructor(
    field: WIProperty,
    paragraphStyles: StyleOptions,
    wiId: number,
    headingLevel: number
  ) {
    this.paragraphStyles = paragraphStyles;
    this.wiId = wiId;
    this.paragraphTemplate = this.generateJsonRichTextParagraphs(
      field,
      this.paragraphStyles
    );
    if (headingLevel && field.name === "Title: ") {
      this.paragraphTemplate.headingLevel = headingLevel;
    } else {
      this.paragraphTemplate.headingLevel = 0;
    }
  } //constructor

  generateJsonRichTextParagraphs(field: WIProperty, paragraphStyles): Run[] {
    let titleStyle = JSON.parse(JSON.stringify(paragraphStyles));

    titleStyle.isBold = true;
    titleStyle.IsUnderline = true;
    titleStyle.InsertSpace = true;
    titleStyle.InsertLineBreak = false;

    let richTextSkin: any[] = [];
    //handle rich text
    try {
      field.richText.forEach((content) => {
        switch (content.type) {
          case "paragraph":
            content.data.fields.forEach((richTextField: WIProperty) => {
              if (/<\/?[a-z][\s\S]*>/.test(richTextField.value)) {
                let html = new JSONHtml(richTextField.value);
                richTextSkin.push(html.getJSONHtml());
              } else {
                //making sure ID is not printed
                let paragraphSkin = new JSONParagraph(
                  richTextField,
                  paragraphStyles,
                  0,
                  0
                );
                richTextSkin.push(paragraphSkin.getJSONParagraph());
              }
            });
            break;
          case "picture":
            richTextSkin.push({
              type: "picture",
              Path: content.data,
            });
            break;
          case "table":
            if (content.data) {
              let tableSkin = new JSONTable(content.data, paragraphStyles, 0);
              richTextSkin.push(tableSkin.getJSONTable());
            }
            break;
        }
      });
      return richTextSkin;
    } catch (e) {
      logger.warn(e);
      return [];
    }
  } //generateJsonRichTextParagraphs

  getJSONRichTextParagraph(): any[] {
    return this.paragraphTemplate;
  } //getJSONTable
} //class
