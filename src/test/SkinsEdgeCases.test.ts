import Skins from '../index';
import logger from '../services/logger';
import JSONRichTextParagraph from '../models/json/paragraph/JSONRichTextParagraph';

jest.mock('../services/logger', () => ({
  __esModule: true,
  default: {
    debug: jest.fn(),
    info: jest.fn(),
    warn: jest.fn(),
    error: jest.fn(),
    silly: jest.fn(),
  },
}));

const getMockLogger = () =>
  logger as unknown as {
    debug: jest.Mock;
    info: jest.Mock;
    warn: jest.Mock;
    error: jest.Mock;
    silly: jest.Mock;
  };

const baseStyles = {
  isBold: false,
  IsItalic: false,
  IsUnderline: false,
  Size: 10,
  Uri: null as string | null,
  Font: 'Arial',
  InsertLineBreak: false,
  InsertSpace: true,
};

const baseHeaderStyles = {
  isBold: true,
  IsItalic: false,
  IsUnderline: false,
  Size: 12,
  Uri: null as string | null,
  Font: 'Arial',
  InsertLineBreak: false,
  InsertSpace: false,
};

describe('Skins – edge cases and public API', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('addNewContentToDocumentSkin throws and logs for unknown skinType', async () => {
    const skins = new Skins('json', 'c://template.dotx');

    await expect(
      skins.addNewContentToDocumentSkin(
        'section-1',
        'unknown-type',
        [],
        baseHeaderStyles as any,
        baseStyles as any,
        0
      )
    ).rejects.toThrow('Unknown skinType : unknown-type - not appended to document skin');

    expect(getMockLogger().error).toHaveBeenCalledWith(
      expect.stringContaining('Fatal error happened when generate skin for : section-1')
    );
  });

  test('generateQueryBasedTable throws and logs for invalid skin format', () => {
    const skins = new Skins('xml', 'c://template.dotx');

    const queryData = [
      {
        url: '',
        level: 0,
        Source: 1,
        fields: [
          { name: 'ID', value: 1, width: '' },
          { name: 'Title', value: 'Item 1', width: '' },
        ],
      },
    ];

    expect(() =>
      skins.generateQueryBasedTable(queryData as any, baseHeaderStyles as any, baseStyles as any, 1)
    ).toThrow('Invalid skin format xml');

    expect(getMockLogger().error).toHaveBeenCalledWith(
      expect.stringContaining('Error occurred in generateQueryBasedTable: Invalid skin format xml')
    );
  });

  test('generateInstallationSkin returns empty array and warns when no data', () => {
    const skins = new Skins('json', 'c://template.dotx');

    const result = skins.generateInstallationSkin([], baseStyles as any);

    expect(result).toEqual([]);
    expect(getMockLogger().warn).toHaveBeenCalledWith('No installation data provided to generate skin');
  });

  test('generateInstallationSkin converts attachments to JSONFile skins', () => {
    const skins = new Skins('json', 'c://template.dotx');

    const installationData = [
      {
        attachment: {
          attachmentLink: 'http://server/file.docx',
          attachmentFileName: 'file.docx',
          attachmentType: 'asEmbedded',
          includeAttachmentContent: true,
        },
      },
      {
        attachment: {
          attachmentLink: 'http://server/image.png',
          attachmentFileName: 'image.png',
          attachmentType: 'asEmbedded',
          includeAttachmentContent: true,
        },
      },
    ];

    const result = skins.generateInstallationSkin(installationData as any, baseStyles as any);

    expect(result).toHaveLength(2);

    expect(result[0]).toMatchObject({
      type: 'File',
      path: 'http://server/file.docx',
      name: 'file',
      isLinkedFile: true,
      includeAttachmentContent: false,
    });

    expect(result[1]).toMatchObject({
      type: 'Picture',
      path: 'http://server/image.png',
      name: 'image',
    });
  });

  test('generateTestReporter returns [] when no data', () => {
    const skins = new Skins('json', 'c://template.dotx');

    const result = skins.generateTestReporter([], 'My Test Plan');
    expect(result).toEqual([]);
  });

  test('generateTestReporter wraps data and plan name', () => {
    const skins = new Skins('json', 'c://template.dotx');
    const data = [{ id: 1, name: 'Suite 1' }];

    const result = skins.generateTestReporter(data as any, 'Plan-1');

    expect(result).toHaveLength(1);
    expect(result[0]).toEqual({
      type: 'testReporter',
      testPlanName: 'Plan-1',
      testSuites: data,
    });
  });

  test('generateTraceTable builds header paragraph and table with grouped header', () => {
    const skins = new Skins('json', 'c://template.dotx');

    const traceData = {
      title: {
        fields: [{ name: 'Title', value: 'Trace Report' }],
        level: 2,
      },
      adoptedData: [
        {
          url: '',
          level: 0,
          Source: 1,
          fields: [
            { name: 'ID', value: 1 },
            { name: 'Name', value: 'Item 1' },
          ],
        },
      ],
      errorMessage: null,
      groupedHeader: {
        leftLabel: 'Requirements',
        rightLabel: 'Tests',
      },
    };

    const result = skins.generateTraceTable(traceData as any, baseHeaderStyles as any, baseStyles as any, 1);

    expect(result).toHaveLength(2);
    expect(result[0]).toMatchObject({ type: 'paragraph' });
    expect(result[1]).toMatchObject({ type: 'table' });

    const table: any = result[1];
    expect(Array.isArray(table.Rows)).toBe(true);
    const groupedRow = table.Rows[0];
    expect(Array.isArray(groupedRow.Cells)).toBe(true);
    expect(groupedRow.Cells[0].gridSpan).toBeGreaterThan(0);
  });

  test('generateTraceTable falls back to error paragraph when no adoptedData but errorMessage exists', () => {
    const skins = new Skins('json', 'c://template.dotx');

    const traceData = {
      title: null,
      adoptedData: [],
      errorMessage: 'Something went wrong',
      groupedHeader: undefined,
    };

    const result = skins.generateTraceTable(traceData as any, baseHeaderStyles as any, baseStyles as any, 0);

    expect(result).toHaveLength(1);
    expect(result[0]).toMatchObject({ type: 'paragraph' });
  });

  test('generateQueryBasedParagraphs handles Description rich text and normal fields', () => {
    const skins = new Skins('json', 'c://template.dotx');

    const data = [
      {
        url: '',
        level: 1,
        Source: 42,
        fields: [
          { name: 'Description: ', value: '<p>Hello</p>' },
          { name: 'Other', value: 'Value' },
        ],
      },
    ];

    const result = skins.generateQueryBasedParagraphs(data as any, baseStyles as any, 1);

    expect(result.length).toBe(2);
    expect(Array.isArray(result[0])).toBe(true);
    expect((result[1] as any).type).toBe('paragraph');
  });

  test('generateQueryBasedParagraphs throws and logs for invalid skin format', () => {
    const skins = new Skins('xml', 'c://template.dotx');

    const data: any[] = [];

    expect(() => skins.generateQueryBasedParagraphs(data as any, baseStyles as any, 0)).toThrow(
      'Invalid skin format xml'
    );

    expect(getMockLogger().error).toHaveBeenCalledWith(
      expect.stringContaining('Error occurred in generateQueryBasedParagraphs: Invalid skin format xml')
    );
  });

  test('validateAndAppendContentControl creates and appends content controls', async () => {
    const skins = new Skins('json', 'c://template.dotx');

    const firstSkin = [{ type: 'paragraph', runs: [] }];
    const secondSkin = [{ type: 'paragraph', runs: [{ text: 'second' }] }];

    const firstResult = await skins.validateAndAppendContentControl('section-1', firstSkin);
    expect(firstResult).toBe(true);

    let documentSkin = skins.getDocumentSkin();
    expect(documentSkin.contentControls).toHaveLength(1);
    expect(documentSkin.contentControls[0].title).toBe('section-1');
    expect(documentSkin.contentControls[0].wordObjects).toEqual(firstSkin);

    const secondResult = await skins.validateAndAppendContentControl('section-1', secondSkin);
    expect(secondResult).toBe(true);

    documentSkin = skins.getDocumentSkin();
    expect(documentSkin.contentControls).toHaveLength(1);
    expect(documentSkin.contentControls[0].wordObjects).toEqual([...firstSkin, ...secondSkin]);
  });

  test('validateAndAppendContentControl logs error and returns undefined on failure', async () => {
    const skins = new Skins('json', 'c://template.dotx');

    const consoleSpy = jest.spyOn(console, 'log').mockImplementation(() => {});

    (skins as any).documentSkin = null;

    const result = await skins.validateAndAppendContentControl('bad-section', []);

    expect(result).toBeUndefined();
    expect(getMockLogger().error).toHaveBeenCalledWith('Error validating and appending to content control');

    consoleSpy.mockRestore();
  });

  test('generateTestBasedSkin throws and logs when suites data is missing', () => {
    const skins = new Skins('json', 'c://template.dotx');

    expect(() =>
      skins.generateTestBasedSkin(null as any, baseHeaderStyles as any, baseStyles as any)
    ).toThrow('No data provided to generate test based skin');

    expect(getMockLogger().error).toHaveBeenCalledWith(
      expect.stringContaining(
        'One or more errors occurred in generateTestBasedSkin: No data provided to generate test based skin'
      )
    );
  });

  test('generateTestBasedSkin throws and logs for invalid skin format', () => {
    const skins = new Skins('xml', 'c://template.dotx');

    const suites: any[] = [];

    expect(() =>
      skins.generateTestBasedSkin(suites as any, baseHeaderStyles as any, baseStyles as any)
    ).toThrow('Invalid skin format xml');

    expect(getMockLogger().error).toHaveBeenCalledWith(
      expect.stringContaining('One or more errors occurred in generateTestBasedSkin: Invalid skin format xml')
    );
  });

  test('generateTestBasedSkin returns empty array when suites array is empty', () => {
    const skins = new Skins('json', 'c://template.dotx');

    const result = skins.generateTestBasedSkin([], baseHeaderStyles as any, baseStyles as any);

    expect(result).toEqual([]);
    expect(getMockLogger().error).not.toHaveBeenCalled();
  });

  test('generateTestBasedSkin inserts page break between single-test suites', () => {
    const skins = new Skins('json', 'c://template.dotx');

    const suites = [
      {
        suiteSkinData: {
          fields: [
            { name: 'Title', value: 'Suite 1' },
            { name: 'ID', value: 1 },
          ],
          level: 1,
        },
        testCases: [
          {
            id: 101,
            testCaseHeaderSkinData: {
              fields: [
                { name: 'Title', value: 'TC-1' },
                { name: 'ID', value: 101 },
                { name: 'Description: ', value: '<p>desc 1</p>' },
              ],
              level: 1,
            },
            testCaseRequirements: [],
            testCaseLinkedMom: [],
            testCaseBugs: [],
            testCaseStepsSkinData: [],
            testCaseAttachments: [],
          },
        ],
      },
      {
        suiteSkinData: {
          fields: [
            { name: 'Title', value: 'Suite 2' },
            { name: 'ID', value: 2 },
          ],
          level: 1,
        },
        testCases: [
          {
            id: 201,
            testCaseHeaderSkinData: {
              fields: [
                { name: 'Title', value: 'TC-2' },
                { name: 'ID', value: 201 },
                { name: 'Description: ', value: '<p>desc 2</p>' },
              ],
              level: 1,
            },
            testCaseRequirements: [],
            testCaseLinkedMom: [],
            testCaseBugs: [],
            testCaseStepsSkinData: [],
            testCaseAttachments: [],
          },
        ],
      },
    ];

    const result = skins.generateTestBasedSkin(suites as any, baseHeaderStyles as any, baseStyles as any, 0);

    const pageBreakParagraphs = result.filter(
      (x: any) => x.type === 'paragraph' && x.runs?.some((r: any) => r.InsertPageBreak)
    );

    expect(pageBreakParagraphs.length).toBe(1);
  });

  test('generateTestBasedSkin in html mode logs info and returns undefined', () => {
    const skins = new Skins('html', 'c://template.dotx');

    const suites = [
      {
        suiteSkinData: {
          fields: [
            { name: 'Title', value: 'Suite' },
            { name: 'ID', value: 1 },
          ],
          level: 1,
        },
        testCases: [],
      },
    ];

    const result = skins.generateTestBasedSkin(suites as any, baseHeaderStyles as any, baseStyles as any, 0);

    expect(result).toBeUndefined();
    expect(getMockLogger().info).toHaveBeenCalledWith('Generating html test data!');
  });

  test('AppendAttachmentContent logs error and records aggregated error when attachment build fails', () => {
    const skins = new Skins('json', 'c://template.dotx');

    const aggregatedErrors: string[] = [];
    const testSkin: any[] = [];

    const data = [
      {
        type: 'SubHeader',
        field: undefined,
      },
    ];

    (skins as any).AppendAttachmentContent(data, testSkin, aggregatedErrors);

    expect(aggregatedErrors.length).toBeGreaterThan(0);
    expect(getMockLogger().error).toHaveBeenCalledWith(
      expect.stringContaining('Error occurred when building attachments')
    );
  });

  test('generateQueryBasedTable in html mode logs info and returns empty array', () => {
    const skins = new Skins('html', 'c://template.dotx');

    const queryData = [
      {
        url: '',
        level: 0,
        Source: 1,
        fields: [
          { name: 'ID', value: 1 },
          { name: 'Title', value: 'Item 1' },
        ],
      },
    ];

    const result = skins.generateQueryBasedTable(
      queryData as any,
      baseHeaderStyles as any,
      baseStyles as any,
      1
    );

    expect(result).toEqual([]);
    expect(getMockLogger().info).toHaveBeenCalledWith('Generating html table!');
  });

  test('generateCoverPageParagraphs in html mode logs info and returns empty array', () => {
    const skins = new Skins('html', 'c://template.dotx');

    const data = [
      {
        url: '',
        level: 0,
        Source: 1,
        fields: [{ name: 'Title', value: 'Cover' }],
      },
    ];

    const result = skins.generateCoverPageParagraphs(data as any, baseStyles as any, 0);

    expect(result).toEqual([]);
    expect(getMockLogger().info).toHaveBeenCalledWith('Generating html cover-page paragraphs!');
  });

  test('generateQueryBasedParagraphs in html mode logs info and returns empty array', () => {
    const skins = new Skins('html', 'c://template.dotx');

    const data = [
      {
        url: '',
        level: 0,
        Source: 1,
        fields: [{ name: 'Title', value: 'Paragraph' }],
      },
    ];

    const result = skins.generateQueryBasedParagraphs(data as any, baseStyles as any, 0);

    expect(result).toEqual([]);
    expect(getMockLogger().info).toHaveBeenCalledWith('Generating html paragraphs!');
  });

  test('generateTraceTable in html mode logs info and returns empty array', () => {
    const skins = new Skins('html', 'c://template.dotx');

    const traceData = {
      title: null,
      adoptedData: [],
      errorMessage: null,
      groupedHeader: undefined,
    };

    const result = skins.generateTraceTable(traceData as any, baseHeaderStyles as any, baseStyles as any, 0);

    expect(result).toEqual([]);
    expect(getMockLogger().info).toHaveBeenCalledWith('Generating html table!');
  });

  test('JSONRichTextParagraph throws for invalid field', () => {
    expect(() => new JSONRichTextParagraph(null as any, baseStyles as any, 1, 0)).toThrow(
      'Invalid field or missing value'
    );

    expect(
      () => new JSONRichTextParagraph({ name: 'Description: ', value: '' } as any, baseStyles as any, 1, 0)
    ).toThrow('Invalid field or missing value');
  });
});
