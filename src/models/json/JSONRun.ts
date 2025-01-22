import { Run, WIProperty, StyleOptions } from './wordJsonModels';
import { defaultJsonRun } from './default';
import striphtml from 'string-strip-html';
import logger from '../../services/logger';

const replaceBr = ({ tag, deleteFrom, deleteTo, rangesArr }) => {
  if (tag.name === `br`) {
    rangesArr.push(deleteFrom, deleteTo, '\n');
  } else {
    rangesArr.push(deleteFrom, deleteTo, ' ');
  }
};

export default class JSONRun {
  runs: Run[];

  constructor(value: string, styles: StyleOptions, fieldtype?: string) {
    this.runs = this.generateJsonRun(value, styles, fieldtype);
  } //constructor

  generateJsonRun(value: string, style: StyleOptions, fieldtype?: string): Run[] {
    let runs: Run[] = [];
    let rowArray;
    try {
      rowArray = value.split('\n');
      //try iterating the array
      rowArray.forEach((text, i) => {
        let run: Run = defaultJsonRun;
        if (fieldtype !== 'SuiteHeaderParagraphTitle') {
          text = `${striphtml(text.toString(), { cb: replaceBr })}` || '';
        }
        if (run.type === 'text') {
          run.value = text;
          run.textStyling.Bold = style.isBold;
          run.textStyling.Italic = style.IsItalic;
          run.textStyling.Underline = style.IsUnderline;
          run.textStyling.Size = style.Size;

          if (run.textStyling.Uri === '') {
            run.textStyling.Uri = null;
          } else {
            run.textStyling.Uri = style.Uri;
          }
          run.textStyling.Font = style.Font;
          if (rowArray.length == 1) {
            run.textStyling.InsertLineBreak = style.InsertLineBreak;
          } else {
            run.textStyling.InsertLineBreak = false;
          }
          if (rowArray.length - 1 == i) {
            run.textStyling.InsertLineBreak = style.InsertLineBreak;
          }
          run.textStyling.InsertSpace = style.InsertSpace;
          runs.push(JSON.parse(JSON.stringify(run)));
        }
      });
    } catch (e) {
      logger.silly(`${value} - cannot be splitted`);
      runs[0] = defaultJsonRun;
      if (defaultJsonRun.type === 'text') {
        runs[0].value = value ? `${striphtml(value.toString(), { cb: replaceBr })}` : ``;
        runs[0].textStyling.Bold = style.isBold;
        runs[0].textStyling.Italic = style.IsItalic;
        runs[0].textStyling.Underline = style.IsUnderline;
        runs[0].textStyling.Size = style.Size;
        if (runs[0].textStyling.Uri === '') {
          runs[0].textStyling.Uri = null;
        } else {
          runs[0].textStyling.Uri = style.Uri;
        }
        runs[0].textStyling.Font = style.Font;
        runs[0].textStyling.InsertLineBreak = style.InsertLineBreak;
        runs[0].textStyling.InsertSpace = style.InsertSpace;
      }
    }

    return runs;
  } //generateJsonRun

  getRun(): Run[] {
    return this.runs;
  }
} //class
