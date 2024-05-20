import { extname } from "path";
import {
  TableCell,
  Run,
  Attachment,
  WIProperty,
  MultipeValuesWIProperty,
  StyleOptions,
  JsonHtml,
} from "../wordJsonModels";
import JSONRun from "../JSONRun";
import logger from "../../../services/logger";

export default class JSONTableCell {
  cell: TableCell;
  constructor(data: WIProperty, tableStyles: StyleOptions) {
    this.cell = this.generateJsonCell(data, tableStyles);
  }

  generateJsonCell(
    data: WIProperty | MultipeValuesWIProperty,
    styles
  ): TableCell {
    let runs: Run[] = [];
    let attachments: Attachment[] = [];
    let HtmlData: string = ""

    //multiple values
    if (
      typeof data.value !== "string" &&
      typeof data.value !== "number" &&
      data.value
    ) {
      data.value.forEach((runData) => {
        //check if photo type
        if (this.isPictureUri(runData.tableCellAttachmentLink)) {
          let attachment = {
            type: "Picture",
            path: runData.tableCellAttachmentLink,
          };
          attachments.push(attachment);
        } 
        else if(/<[^>]*>/.test(runData.value)) {
          HtmlData = runData.value || "";
        }
        else {
          styles.Uri = runData.relativeUrl
            ? runData.relativeUrl
            : runData.Uri
            ? runData.Uri
            : null;
          styles.InsertLineBreak = false;
          let jsonRun = new JSONRun(runData.value, styles);
          runs = [...runs, ...jsonRun.getRun()];
        }
      });
    }
    else if(/<[^>]*>/.test(data.value)) {
      HtmlData = data.value || "";
    }
     else {
      styles.Uri = data.relativeUrl
        ? data.relativeUrl
        : data.url
        ? data.url
        : null;
      let text = data.value || "";
      let jsonRun = new JSONRun(`${text}`, styles);
      runs = [...runs, ...jsonRun.getRun()];
    }
    if (HtmlData!= "")
    {
      let Html:JsonHtml = {
        type:"html",
        Html:HtmlData
      }
      return {
        attachments: attachments,
        Paragraphs: [
          {
            Runs: runs,
          },
        ],
        Html
      };
    }
    else{
      return {
        attachments: attachments,
      Paragraphs: [
        {
          Runs: runs,
        },
      ],
    };
  }
} //generateJsonCells

  isPictureUri(uri: string) {
    try {
      let fileExt = extname(uri);
      logger.debug(`attachment extension: ${fileExt}`);
      if (fileExt == ".PNG" || fileExt == ".jpg") {
        return true;
      } else {
        return false;
      }
    } catch (e) {
      logger.error(`error parsing file extension for : ${uri}`);
      return false;
    }
  }

  getJsonCell(): TableCell {
    return this.cell;
  }
} //class
