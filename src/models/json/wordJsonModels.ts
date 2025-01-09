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
  font: string;
  fontSize: number;
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
  shading?: Shading;
  width?: string;
  Html?: JsonHtml;
}

export interface Shading {
  color: string;
  fill: string;
  themeFillShade?: string;
}

export interface Paragraph {
  type: string;
  headingLevel?: number;
  runs: Run[];
}
export interface Attachment {
  type: string;
  path: string;
  name: string;
  isLinkedFile?: boolean;
  IsFlattened?: boolean;
}

interface TextStyle {
  Bold: boolean;
  Italic: boolean;
  Underline: boolean;
  Size: number;
  Uri: string;
  Font: string;
  InsertLineBreak: boolean;
  InsertSpace: boolean;
  src?: string; //for image
}

export interface Run {
  type: 'text' | 'image' | 'break' | 'other';
  value?: string;
  src?: string; // If type = 'image', you can store the URL or local path here
  textStyling?: TextStyle; // If type = 'text', store style; if type = 'image', optional or partial
}
export interface WorkItemData {}

export interface Table {
  type: string;
  headingLevel: number;
  Rows: TableRow[];
  insertPageBreak: boolean;
}

export interface List {
  type: string;
  listItems: ListItem[];
  isOrdered: boolean;
}

export interface ListItem {
  Runs: Run[];
  level: number;
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
  width?: string;
  shading?: Shading;
  url?: string;
  relativeUrl?: string;
  attachmentLink?: string;
  attachmentType?: string;
  relativeAttachmentLink?: string;
  richTextNodes?: any[];
}
export interface MultipeValuesWIProperty {
  name: string;
  value: any[];
  width?: string;
  shading?: Shading;
  url?: string;
  relativeUrl?: string;
  attachmentLink?: string;
  attachmentType?: string;
  relativeAttachmentLink?: string;
  richTextNodes?: any[];
}
export interface WIColumns {
  name: string;
  referenceName: string;
  url: string;
}

export type RichNode = ParagraphNode | TextNode | ImageNode | TableNode | ListNode | BreakNode | OtherNode;

export interface ParagraphNode {
  type: 'paragraph';
  children: RichNode[];
}

export interface TextNode {
  type: 'text';
  value: string;
}

export interface ImageNode {
  type: 'image';
  src: string;
  alt?: string;
}

export interface BreakNode {
  type: 'break';
}

export interface TableNode {
  type: 'table';
  children: RichNode[];
}

export interface ListNode {
  type: 'list';
  isOrdered: boolean;
  children: RichNode[];
}

export interface OtherNode {
  type: 'other';
  tagName: string;
  children: RichNode[];
}
