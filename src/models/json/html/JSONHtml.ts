import { JsonHtml } from "../wordJsonModels";
import JSONRun from "../JSONRun";

export default class JSONHtml {
  htmlString: string = "";
  constructor(htmlString: string) {
    this.htmlString = htmlString;
  } //constructor

  getJSONHtml(): JsonHtml {
    return { type: "html", Html: this.htmlString };
  } //getJSONTable
} //class
