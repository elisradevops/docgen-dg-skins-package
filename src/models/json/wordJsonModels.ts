export interface StyleOptions {
  isBold: boolean;
  IsItalic: boolean;
  IsUnderline: boolean;
  Size: number;
  Uri: string;
  Font: string;
  InsertLineBreak: boolean;
  InsertSpace: boolean;
}
export interface DocumentSkin {
  templatePath: string;
  contentControls: ContetControlSkin[];
}
export interface ContetControlSkin {
  title: string;
  wordObjects: any[]; //can be table or paragraph or file
}
export interface JsonHtml {
  type: string;
  Html: string;
}
export interface TableRow {
  Cells: TableCell[];
}
export interface TableCell {
  attachments?: Attachment[];
  Paragraphs: [
    {
      Runs: Run[];
    }
  ];
  Html?:JsonHtml
}
export interface Paragraph {
  type: string;
  headingLevel?: number;
  runs: Run[];
}
export interface Attachment {
  type: string;
  path: string;
}
export interface Run {
  text: string;
  Bold: boolean;
  Italic: boolean;
  Underline: boolean;
  Size: number;
  Uri: string;
  Font: string;
  InsertLineBreak: boolean;
  InsertSpace: boolean;
}
export interface WorkItemData {}

export interface Table {
  type: string;
  headingLevel: number;
  Rows: TableRow[];
}
export interface WIQueryResults {
  columns: WIColumns[];
  workItems: WIData[];
  asOf: string;
  queryResultType: number;
  queryType: number;
}
export interface WIData {
  url: string;
  fields: WIProperty[];
  Source: number;
  level: number;
}
export interface WIProperty {
  name: string;
  value: any;
  url?: string;
  relativeUrl?: string;
  attachmentLink?: string;
  relativeAttachmentLink?: string;
  richText?: any[];
}
export interface MultipeValuesWIProperty {
  name: string;
  value: any[];
  url?: string;
  relativeUrl?: string;
  attachmentLink?: string;
  relativeAttachmentLink?: string;
  richText?: any[];
}
export interface WIColumns {
  name: string;
  referenceName: string;
  url: string;
}
