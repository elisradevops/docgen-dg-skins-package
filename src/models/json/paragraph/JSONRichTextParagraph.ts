import { Run, WIProperty, StyleOptions } from '../wordJsonModels';
import JSONHtml from '../html/JSONHtml';
import logger from '../../../services/logger';

export default class JSONRichTextParagraph {
  paragraphStyles: StyleOptions;
  paragraphTemplate: any;
  wiId: number;
  constructor(
    field: WIProperty,
    paragraphStyles: StyleOptions,
    wiId: number,
    headingLevel: number,
    skipImageFormat = false
  ) {
    if (!field || !field.value) {
      throw new Error('Invalid field or missing value');
    }

    this.paragraphStyles = paragraphStyles || {
      isBold: false,
      IsItalic: false,
      IsUnderline: false,
      Size: 12,
      Uri: null,
      Font: 'Arial',
      InsertLineBreak: false,
      InsertSpace: false,
    };
    this.wiId = wiId;
    this.paragraphTemplate = this.generateJsonRichTextParagraphs(
      field,
      this.paragraphStyles,
      skipImageFormat
    );
    if (headingLevel && field.name === 'Title: ') {
      this.paragraphTemplate.headingLevel = headingLevel;
    }
  } //constructor

  generateJsonRichTextParagraphs(field: WIProperty, paragraphStyles, skipImageFormat = false): Run[] {
    let titleStyle = JSON.parse(JSON.stringify(paragraphStyles));

    titleStyle.isBold = true;
    titleStyle.IsUnderline = true;
    titleStyle.InsertSpace = true;
    titleStyle.InsertLineBreak = false;

    try {
      let richTextSkin: any[] = [];
      let html = new JSONHtml(field.value, paragraphStyles.Font, paragraphStyles.Size);
      richTextSkin.push(html.getJSONHtml());
      return richTextSkin;
    } catch (e) {
      logger.error(`Error parsing rich text from JSONRichTextParagraph: ${e.message}`);
      return [];
    }
  } //generateJsonRichTextParagraphs

  getJSONRichTextParagraph(): any[] {
    return this.paragraphTemplate;
  } //getJSONRichTextParagraph
} //class
