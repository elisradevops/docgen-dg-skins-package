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

// fieldtypes that bypass striphtml entirely — used for text that must survive byte-for-byte
// (e.g. a deliberate single-space spacer run, which striphtml would otherwise trim to '').
const RAW_TEXT_FIELD_TYPES = ['SuiteHeaderParagraphTitle', 'RawText'];

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
        if (!RAW_TEXT_FIELD_TYPES.includes(fieldtype)) {
          // string-strip-html unconditionally trims trailing whitespace — this silently ate the space
          // JSONParagraph appends after a field label (e.g. "Called Date: " -> "Called Date:"),
          // producing "Called Date:07/12/2026" with no space before the value. Restore it when the
          // original text actually had trailing whitespace (never invent a space that wasn't there).
          const original = text.toString();
          const hadTrailingSpace = /\s$/.test(original);
          const stripped = `${striphtml(original, { cb: replaceBr })}`;
          text = hadTrailingSpace && stripped && !/\s$/.test(stripped) ? `${stripped} ` : stripped || '';
        }
        run.text = text;
        run.Bold = style.isBold;
        run.Italic = style.IsItalic;
        run.Underline = style.IsUnderline;
        run.Size = style.Size;
        run.InsertPageBreak = style.InsertPageBreak;
        if (run.Uri === '') {
          run.Uri = null;
        } else {
          run.Uri = style.Uri;
        }
        run.Font = style.Font;
        if (rowArray.length == 1) {
          run.InsertLineBreak = style.InsertLineBreak;
        } else if (rowArray.length > 1 && i < rowArray.length - 1) {
          run.InsertLineBreak = rowArray.length > 1;
        } else if (rowArray.length - 1 == i) {
          run.InsertLineBreak = style.InsertLineBreak;
        }
        run.InsertSpace = style.InsertSpace;
        run.SectionRefId = style.SectionRefId || null;
        runs.push(JSON.parse(JSON.stringify(run)));
      });
    } catch (e) {
      logger.silly(`${value} - cannot be splitted`);
      runs[0] = defaultJsonRun;
      runs[0].text = value ? `${striphtml(value.toString(), { cb: replaceBr })}` : ``;
      runs[0].Bold = style.isBold;
      runs[0].Italic = style.IsItalic;
      runs[0].Underline = style.IsUnderline;
      runs[0].Size = style.Size;
      if (runs[0].Uri === '') {
        runs[0].Uri = null;
      } else {
        runs[0].Uri = style.Uri;
      }
      runs[0].Font = style.Font;
      runs[0].InsertLineBreak = style.InsertLineBreak;
      runs[0].InsertSpace = style.InsertSpace;
      runs[0].SectionRefId = style.SectionRefId || null;
    }

    return runs;
  } //generateJsonRun

  getRun(): Run[] {
    return this.runs;
  }
} //class
