import { JsonHtml } from '../wordJsonModels';
import JSONRun from '../JSONRun';

export default class JSONHtml {
  htmlString: string = '';
  font: string = '';
  fontSize: number = 0;
  constructor(htmlString: string, font = 'Arial', fontSize = 12) {
    this.htmlString = htmlString;
    this.font = font;
    this.fontSize = fontSize;
  } //constructor

  getJSONHtml(): JsonHtml {
    return { type: 'html', Html: this.htmlString, font: this.font, fontSize: this.fontSize };
  } //getJSONTable
} //class
