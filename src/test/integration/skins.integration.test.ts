import { writeFileSync } from 'fs';
import Skins from '../../index';

jest.mock('../../services/logger', () => ({
  __esModule: true,
  default: {
    debug: jest.fn(),
    info: jest.fn(),
    warn: jest.fn(),
    error: jest.fn(),
    silly: jest.fn(),
  },
}));

jest.setTimeout(300000);
const styles = {
  isBold: false,
  IsItalic: false,
  IsUnderline: false,
  Size: 10,
  Uri: null,
  Font: 'New Times Roman',
  InsertLineBreak: true,
  InsertSpace: true,
};

const headerStyles = {
  isBold: true,
  IsItalic: false,
  IsUnderline: false,
  Size: 12,
  Uri: null,
  Font: 'Arial',
  InsertLineBreak: false,
  InsertSpace: true,
};

// Minimal inline fixtures to replace removed JSON mock-data and snapshot files
const flatQueryData = [
  {
    url: 'http://example.com/wi/1',
    level: 0,
    Source: 1,
    fields: [
      { name: 'ID', value: 1, width: '50%' },
      { name: 'Title', value: 'First item', width: '50%' },
    ],
  },
  {
    url: 'http://example.com/wi/2',
    level: 0,
    Source: 2,
    fields: [
      { name: 'ID', value: 2, width: '50%' },
      { name: 'Title', value: 'Second item', width: '50%' },
    ],
  },
];

const treeQueryData = [
  {
    url: 'http://example.com/wi/10',
    level: 0,
    Source: 10,
    fields: [
      { name: 'ID', value: 10 },
      { name: 'Title', value: 'Parent item' },
    ],
  },
  {
    url: 'http://example.com/wi/11',
    level: 1,
    Source: 11,
    fields: [
      { name: 'ID', value: 11 },
      { name: 'Title', value: 'Child item' },
    ],
  },
];

const simpleTestSuites = [
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
            { name: 'Description: ', value: '<p>Simple description</p>' },
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

const complexTestSuites = [
  {
    suiteSkinData: {
      fields: [
        { name: 'Title', value: 'Complex Suite' },
        { name: 'ID', value: 200 },
      ],
      level: 1,
    },
    testCases: [
      {
        id: 201,
        testCaseHeaderSkinData: {
          fields: [
            { name: 'Title', value: 'TC-Complex' },
            { name: 'ID', value: 201 },
            { name: 'Description: ', value: '<p>Complex <b>HTML</b> description</p>' },
          ],
          level: 1,
        },
        testCaseRequirements: [
          {
            url: 'http://example.com/req/1',
            level: 0,
            Source: 1,
            fields: [
              { name: 'Requirement', value: 'REQ-1' },
              { name: 'Title', value: 'First requirement' },
            ],
          },
        ],
        testCaseLinkedMom: [
          {
            url: 'http://example.com/mom/1',
            level: 0,
            Source: 2,
            fields: [
              { name: 'MOM', value: 'MOM-1' },
              { name: 'Title', value: 'First MOM' },
            ],
          },
        ],
        testCaseBugs: [
          {
            url: 'http://example.com/bug/1',
            level: 0,
            Source: 3,
            fields: [
              { name: 'Bug', value: 'BUG-1' },
              { name: 'Title', value: 'First bug' },
            ],
          },
        ],
        testCaseStepsSkinData: [
          {
            url: 'http://example.com/step/1',
            level: 0,
            Source: 4,
            fields: [
              { name: 'Step', value: 'Do something' },
              { name: 'Expected', value: 'It works' },
            ],
          },
        ],
        testCaseAttachments: [
          {
            url: 'http://example.com/attach/1',
            level: 0,
            Source: 5,
            fields: [{ name: 'Attachment', value: 'log.txt' }],
          },
        ],
        testCaseDocAttachmentsAdoptedData: {
          stepLevel: [
            {
              type: 'SubHeader',
              field: { name: 'Title', value: 'Step Attachments' },
            },
            {
              type: 'File',
              attachmentLink: 'http://example.com/doc1.pdf',
              attachmentFileName: 'doc1.pdf',
              attachmentType: 'asEmbedded',
              includeAttachmentContent: true,
            },
          ],
          testCaseLevel: [
            {
              type: 'SubHeader',
              field: { name: 'Title', value: 'Case Attachments' },
            },
            {
              type: 'File',
              attachmentLink: 'http://example.com/doc2.pdf',
              attachmentFileName: 'doc2.pdf',
              attachmentType: 'asEmbedded',
              includeAttachmentContent: true,
            },
          ],
        },
        insertPageBreak: true,
      },
    ],
  },
];

const changesData = [
  {
    artifact: [
      {
        url: 'http://example.com/artifact/1',
        level: 0,
        Source: 1,
        fields: [
          { name: 'Title', value: 'Artifact 1' },
          { name: 'Description: ', value: '<p>Artifact description</p>' },
        ],
      },
    ],
    artifactChanges: [
      {
        url: 'http://example.com/change/1',
        level: 0,
        Source: 2,
        fields: [
          { name: 'Field', value: 'Status' },
          { name: 'Old', value: 'Open' },
          { name: 'New', value: 'Closed' },
        ],
      },
    ],
  },
];

const richTextTestSuites = [
  {
    suiteSkinData: {
      fields: [
        { name: 'Title', value: 'Rich Suite' },
        { name: 'ID', value: 300 },
      ],
      level: 1,
    },
    testCases: [
      {
        id: 301,
        testCaseHeaderSkinData: {
          fields: [
            { name: 'Title', value: 'TC-Rich' },
            { name: 'ID', value: 301 },
            { name: 'Description: ', value: '<p><b>Rich</b> text description</p>' },
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

describe('Generate json skins from queries - tests', () => {
  test('generate table skin - flat query', async () => {
    let skins = new Skins('json', 'c:\test.dotx');
    const result = await skins.addNewContentToDocumentSkin(
      'test',
      skins.SKIN_TYPE_TABLE,
      flatQueryData as any,
      headerStyles,
      styles,
      4
    );

    expect(Array.isArray(result)).toBe(true);
    expect(result.length).toBeGreaterThan(0);
    const table = result[0] as any;
    expect(table.type).toBe('table');
    expect(Array.isArray(table.Rows)).toBe(true);
    expect(table.Rows.length).toBeGreaterThanOrEqual(3); // grouped/header + header + data
  });
  test('generate table skin - tree query', async () => {
    let skins = new Skins('json', 'c:\test.dotx');
    const result = await skins.addNewContentToDocumentSkin(
      'test',
      skins.SKIN_TYPE_TABLE,
      treeQueryData as any,
      headerStyles,
      styles,
      4
    );

    const table = result[0] as any;
    expect(table.type).toBe('table');
    expect(table.headingLevel).toBe(4);
    expect(table.Rows.length).toBeGreaterThanOrEqual(3);
  });
  test('generate paragraph skin - flat query', async () => {
    let skins = new Skins('json', 'c:\test.dotx');
    const result = await skins.addNewContentToDocumentSkin(
      'test',
      skins.SKIN_TYPE_PARAGRAPH,
      flatQueryData as any,
      headerStyles,
      styles,
      4
    );
    // writeFileSync(
    //   "flat-query-paragraph-snapshot.json",
    //   JSON.stringify(documentSkin)
    // );

    expect(Array.isArray(result)).toBe(true);
    expect(result.length).toBeGreaterThan(0);
    result.forEach((p: any) => {
      expect(p.type).toBe('paragraph');
    });
  });
  test('generate paragraph skin - tree query', async () => {
    let skins = new Skins('json', 'c:\test.dotx');
    const result = await skins.addNewContentToDocumentSkin(
      'test',
      skins.SKIN_TYPE_PARAGRAPH,
      treeQueryData as any,
      headerStyles,
      styles,
      4
    );
    // writeFileSync("tree-query-data.json", JSON.stringify(documentSkin));

    expect(Array.isArray(result)).toBe(true);
    expect(result.length).toBeGreaterThan(0);
    result.forEach((p: any) => {
      expect(p.type).toBe('paragraph');
    });
  });
});
describe('Generate json skins from testData - tests', () => {
  test('generate std skin - testData', async () => {
    let skins = new Skins('json', 'c\\test\\test.dotx');
    const result = await skins.addNewContentToDocumentSkin(
      'system-capabilities',
      skins.SKIN_TYPE_TEST_PLAN,
      simpleTestSuites as any,
      headerStyles,
      styles,
      4
    );

    expect(Array.isArray(result)).toBe(true);
    expect(result.length).toBeGreaterThan(0);
    // Expect suite header + test case header + description title + rich text/html
    const hasParagraph = result.some((x: any) => x.type === 'paragraph');
    expect(hasParagraph).toBe(true);
  });
  test('generate std skin - complex testData', async () => {
    let skins = new Skins('json', 'C\\docgen\\documents\\181020205911\\test.dotx');
    const result = await skins.addNewContentToDocumentSkin(
      'system-capabilities',
      skins.SKIN_TYPE_TEST_PLAN,
      complexTestSuites as any,
      headerStyles,
      styles,
      4
    );

    expect(Array.isArray(result)).toBe(true);
    expect(result.length).toBeGreaterThan(0);
    const hasPageBreak = result.some(
      (x: any) => x.type === 'paragraph' && x.runs?.some((r: any) => r.InsertPageBreak)
    );
    expect(hasPageBreak).toBe(true);
  });
  test('generate test-std skin alias', async () => {
    let skins = new Skins('json', 'c\\test\\test.dotx');
    const result = await skins.addNewContentToDocumentSkin(
      'system-capabilities',
      skins.SKIN_TYPE_TEST_STD,
      simpleTestSuites as any,
      headerStyles,
      styles,
      4
    );

    expect(Array.isArray(result)).toBe(true);
    expect(result.length).toBeGreaterThan(0);
  });
  test('generate test-stp skin alias', async () => {
    let skins = new Skins('json', 'c\\test\\test.dotx');
    const stpSuitesWithPhase = [
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
                { name: 'Description: ', value: '<p>Simple description</p>' },
                { name: 'Test Phase', value: 'Qualification' },
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
    const result = await skins.addNewContentToDocumentSkin(
      'system-capabilities',
      skins.SKIN_TYPE_TEST_STP,
      stpSuitesWithPhase as any,
      headerStyles,
      styles,
      4
    );

    expect(Array.isArray(result)).toBe(true);
    expect(result.length).toBeGreaterThan(0);
    const hasPhaseParagraph = result.some((paragraph: any) => {
      if (paragraph?.type !== 'paragraph' || !Array.isArray(paragraph?.runs)) return false;
      const text = paragraph.runs.map((run: any) => run?.text || '').join('');
      return text.includes('Test Phase: Qualification');
    });
    expect(hasPhaseParagraph).toBe(true);
  });
  test('generate test-stp skin alias with +1 heading depth vs std', async () => {
    let skins = new Skins('json', 'c\\test\\test.dotx');
    const suites = [
      {
        suiteSkinData: {
          fields: [
            { name: 'Title', value: 'Suite H' },
            { name: 'ID', value: 11 },
          ],
          level: 1,
        },
        testCases: [
          {
            id: 111,
            testCaseHeaderSkinData: {
              fields: [
                { name: 'Title', value: 'TC-H' },
                { name: 'ID', value: 111 },
                { name: 'Description: ', value: '<p>Simple description</p>' },
              ],
              level: 2,
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

    const stdResult = await skins.addNewContentToDocumentSkin(
      'system-capabilities',
      skins.SKIN_TYPE_TEST_STD,
      suites as any,
      headerStyles,
      styles,
      4
    );
    const stpResult = await skins.addNewContentToDocumentSkin(
      'system-capabilities',
      skins.SKIN_TYPE_TEST_STP,
      suites as any,
      headerStyles,
      styles,
      4
    );

    const getHeadingLevelByText = (items: any[], text: string) => {
      const paragraph = items.find((item) => {
        if (item?.type !== 'paragraph' || !Array.isArray(item?.runs)) return false;
        const runText = item.runs.map((run: any) => String(run?.text || '')).join('');
        return runText.includes(text) && Number.isInteger(item?.headingLevel) && item.headingLevel > 0;
      });
      return paragraph?.headingLevel;
    };

    const stdSuiteHeading = getHeadingLevelByText(stdResult as any[], 'Suite H');
    const stpSuiteHeading = getHeadingLevelByText(stpResult as any[], 'Suite H');
    const stdCaseHeading = getHeadingLevelByText(stdResult as any[], 'TC-H');
    const stpCaseHeading = getHeadingLevelByText(stpResult as any[], 'TC-H');

    expect(stpSuiteHeading).toBe(stdSuiteHeading + 1);
    expect(stpCaseHeading).toBe(stdCaseHeading + 1);
  });
  test('generate test-str skin alias', async () => {
    let skins = new Skins('json', 'c\\test\\test.dotx');
    const result = await skins.addNewContentToDocumentSkin(
      'appendix-a-content-control',
      skins.SKIN_TYPE_TEST_STR,
      [
        {
          type: 'Header',
          field: { name: 'Title', value: 'Section A' },
          level: 1,
        },
        {
          field: { name: 'Description', value: 'Row 1' },
        },
      ] as any,
      headerStyles,
      styles,
      4
    );

    expect(Array.isArray(result)).toBe(true);
    expect(result.length).toBeGreaterThan(0);
  });
  test('generate trace table skin', async () => {
    let skins = new Skins('json', 'c\\test\\test.dotx');
    const traceData = {
      title: {
        fields: [{ name: 'Title', value: 'Trace Table' }],
        level: 1,
      },
      adoptedData: flatQueryData,
      errorMessage: null,
      groupedHeader: {
        leftLabel: 'Left',
        rightLabel: 'Right',
      },
    };

    const result = await skins.addNewContentToDocumentSkin(
      'system-capabilities',
      skins.SKIN_TYPE_TRACE,
      traceData as any,
      headerStyles,
      styles,
      4
    );
    expect(result.length).toBe(2);
    expect(result[0]).toMatchObject({ type: 'paragraph' });
    expect(result[1]).toMatchObject({ type: 'table' });
  });
});
describe('Generate json skins svd - tests', () => {
  test('generate changes between releases skin', async () => {
    let skins = new Skins('json', 'c\\test\\test.dotx');
    for (let i = 0; i < changesData.length; i++) {
      let artifactChangesData = changesData[i];
      const paragraphSkin = await skins.addNewContentToDocumentSkin(
        'change-description-content-control',
        skins.SKIN_TYPE_PARAGRAPH,
        artifactChangesData.artifact,
        headerStyles,
        styles,
        4
      );

      const tableSkin = await skins.addNewContentToDocumentSkin(
        'change-description-content-control',
        skins.SKIN_TYPE_TABLE,
        artifactChangesData.artifactChanges,
        headerStyles,
        styles,
        4
      );
      expect(paragraphSkin.length).toBeGreaterThan(0);
      expect(tableSkin[0]).toMatchObject({ type: 'table' });
    }
  });
});
describe('Common functions - tests', () => {
  test('generate richText skin', async () => {
    let skins = new Skins('json', 'c\\test\\test.dotx');
    const result = await skins.addNewContentToDocumentSkin(
      'system-capabilities',
      skins.SKIN_TYPE_TEST_PLAN,
      richTextTestSuites as any,
      headerStyles,
      styles,
      4
    );
    const hasHtml = result.some((x: any) => x.type === 'html');
    expect(hasHtml).toBe(true);
  });
});

describe('Time machine report skin - tests', () => {
  test('generate time-machine-report skin with compare summary and differences', async () => {
    const skins = new Skins('json', 'c\\test\\test.dotx');
    const result = await skins.addNewContentToDocumentSkin(
      'historical-compare-report-content-control',
      skins.SKIN_TYPE_TIME_MACHINE,
      {
        teamProjectName: 'MEWP',
        queryName: 'Shared Query',
        compareResult: {
          baseline: { asOf: '2025-12-22T17:08:00.000Z', total: 4 },
          compareTo: { asOf: '2025-12-28T08:57:00.000Z', total: 4 },
          summary: { updatedCount: 1 },
          rows: [
            {
              id: 11,
              workItemType: 'Requirement',
              title: 'Req-11',
              workItemUrl: 'https://dev.azure.com/org/project/_workitems/edit/11',
              baselineRevisionId: 2,
              compareToRevisionId: 20,
              compareStatus: 'Changed',
              differences: [{ field: 'Test Phase', baseline: 'FAT', compareTo: 'FAT; ATP' }],
            },
          ],
        },
      },
      headerStyles,
      styles,
      1,
    );

    expect(Array.isArray(result)).toBe(true);
    expect(result.length).toBeGreaterThan(0);
    expect(result.some((item: any) => item.type === 'table')).toBe(true);
    expect(
      result.some((item: any) => {
        if (item?.type !== 'paragraph') return false;
        const text = (item?.runs || []).map((run: any) => run?.text || '').join('');
        return text.includes('Difference');
      }),
    ).toBe(true);
    expect(
      result.some((item: any) => {
        if (item?.type !== 'paragraph') return false;
        const text = (item?.runs || []).map((run: any) => run?.text || '').join('');
        return text.includes('Work Item Differences');
      }),
    ).toBe(true);
  });
});
