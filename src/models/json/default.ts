import { Run, StyleOptions } from './wordJsonModels';

export let defaultJsonRun: Run = {
  type: 'text',
  value: ``,
  textStyling: {
    Bold: false,
    Italic: false,
    Underline: false,
    Size: 12,
    Uri: null,
    Font: 'Arial',
    InsertLineBreak: false,
    InsertSpace: true,
  },
};

export let DescriptionandProcedureStyle: StyleOptions = {
  isBold: true,
  IsItalic: false,
  IsUnderline: true,
  Size: 12,
  Uri: null,
  Font: 'Arial',
  InsertLineBreak: false,
  InsertSpace: false,
};
