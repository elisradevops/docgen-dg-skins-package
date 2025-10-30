export interface StyleOptions {
  isBold: boolean;
  IsItalic: boolean;
  IsUnderline: boolean;
  Size: number;
  Uri: string;
  Font: string;
  InsertLineBreak: boolean;
  InsertSpace: boolean;
  InsertPageBreak?: boolean;
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
  vMerge?: 'restart' | 'continue';
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
  includeAttachmentContent?: boolean;
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
  InsertPageBreak?: boolean;
  src?: string; //for image
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

export interface TestSuite {
  suiteName: string;
  testCases: TestCase[];
}

export interface TestCase {
  testCaseId: number;
  testCaseName: string;
  testSteps?: TestStep[];
  testCaseResult: string;
  priority?: number;
  assignedTo?: string;
  subSystem?: string;
  runBy?: string;
  configuration?: string;
  automationStatus?: string;
  associatedRequirement?: string;
  associatedBug?: string;
  associatedCR?: string;
}

export interface TestStep {
  stepNo: number;
  stepAction: string;
  stepExpected: string;
  stepStatus: string;
  stepComments: string;
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
  includeAttachmentContent?: boolean;
  relativeAttachmentLink?: string;
  richText?: any[];
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
  includeAttachmentContent?: boolean;
  relativeAttachmentLink?: string;
  richText?: any[];
}
export interface WIColumns {
  name: string;
  referenceName: string;
  url: string;
}
