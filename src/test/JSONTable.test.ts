import JSONTable from '../models/json/table/JSONTable';

const baseStyles = {
  isBold: false,
  IsItalic: false,
  IsUnderline: false,
  Size: 10,
  Uri: null as string | null,
  Font: 'Arial',
  InsertLineBreak: false,
  InsertSpace: false,
};

const headerStyles = {
  ...baseStyles,
  isBold: true,
};

const buildData = (secondColValue: string) => [
  {
    url: '',
    Source: 1,
    level: 1,
    fields: [
      { name: '#', value: 1, width: '8%' },
      { name: 'Description', value: 'No description' },
    ],
  },
  {
    url: '',
    Source: 2,
    level: 1,
    fields: [
      { name: '#', value: 2, width: '8%' },
      { name: 'Description', value: secondColValue },
    ],
  },
];

describe('JSONTable vertical merge behavior', () => {
  test('uses explicit grouped-header column spans when provided', () => {
    const table = new JSONTable(
      [
        {
          url: '',
          Source: 1,
          level: 1,
          fields: [
            { name: 'Req ID', value: 1 },
            { name: 'Title', value: 'Req' },
            { name: 'Customer ID', value: 'C1' },
            { name: 'Test Case ID', value: 11 },
            { name: 'Title', value: 'TC' },
          ],
        },
      ] as any,
      headerStyles as any,
      baseStyles as any,
      0,
      false,
      false,
      false,
      {
        leftLabel: 'Requirement',
        rightLabel: 'Test Case',
        leftColumns: 3,
        rightColumns: 2,
      },
      false
    ).getJSONTable() as any;

    expect(table.Rows[0].Cells[0].gridSpan).toBe(3);
    expect(table.Rows[0].Cells[1].gridSpan).toBe(2);
  });

  test('does not merge empty cells when vertical merge is disabled', () => {
    const table = new JSONTable(
      buildData('') as any,
      headerStyles as any,
      baseStyles as any,
      0,
      false,
      false,
      false,
      undefined,
      false
    ).getJSONTable() as any;

    expect(table.Rows[1].Cells[1].vMerge).toBeUndefined();
    expect(table.Rows[2].Cells[1].vMerge).toBeUndefined();
  });

  test('merges empty cells when vertical merge is enabled', () => {
    const table = new JSONTable(
      buildData('') as any,
      headerStyles as any,
      baseStyles as any,
      0,
      false,
      false,
      false,
      undefined,
      true
    ).getJSONTable() as any;

    expect(table.Rows[1].Cells[1].vMerge).toBe('restart');
    expect(table.Rows[2].Cells[1].vMerge).toBe('continue');
  });

  test('does not treat HTML cell content as empty during merge detection', () => {
    const table = new JSONTable(
      buildData('<div>Some Description for suite</div>') as any,
      headerStyles as any,
      baseStyles as any,
      0,
      false,
      false,
      false,
      undefined,
      true
    ).getJSONTable() as any;

    expect(table.Rows[2].Cells[1].vMerge).toBeUndefined();
    expect(table.Rows[2].Cells[1].Html?.Html).toBe('<div>Some Description for suite</div>');
  });
});

describe('JSONTable empty query results', () => {
  test('renders a single informational row instead of throwing when the query returns zero rows', () => {
    const table = new JSONTable([] as any, headerStyles as any, baseStyles as any, 0).getJSONTable() as any;

    expect(table.Rows).toHaveLength(1);
    expect(table.Rows[0].Cells).toHaveLength(1);
    expect(table.Rows[0].Cells[0].Paragraphs[0].Runs[0].text).toBe('No results found for this query.');
  });

  test('renders the same informational row when data is undefined/null', () => {
    const table = new JSONTable(undefined as any, headerStyles as any, baseStyles as any, 0).getJSONTable() as any;

    expect(table.Rows).toHaveLength(1);
    expect(table.Rows[0].Cells[0].Paragraphs[0].Runs[0].text).toBe('No results found for this query.');
  });
});

describe('JSONTable adaptive layout (enableAdaptiveLayout)', () => {
  const buildAdaptiveData = () => [
    {
      url: '',
      Source: 1,
      level: 1,
      fields: [
        { name: 'ID', value: 1327 },
        { name: 'Stack Rank', value: '' },
        { name: 'Title', value: 'Requirement B Subject 2.2' },
      ],
    },
    {
      url: '',
      Source: 2,
      level: 1,
      fields: [
        { name: 'ID', value: 1155 },
        { name: 'Stack Rank', value: '' },
        { name: 'Title', value: 'Some bug to fix' },
      ],
    },
  ];

  const buildTable = (data: any, enableAdaptiveLayout: boolean) =>
    new JSONTable(
      data,
      headerStyles as any,
      baseStyles as any,
      0,
      false,
      false,
      false,
      undefined,
      false,
      enableAdaptiveLayout
    ).getJSONTable() as any;

  test('leaves columns and widths untouched when adaptive layout is disabled (default)', () => {
    const table = buildTable(buildAdaptiveData(), false);

    // 3 original columns survive, including the all-empty Stack Rank column
    expect(table.Rows[0].Cells).toHaveLength(3);
    expect(table.Rows[0].Cells[1].width).toBe('');
  });

  test('drops a column that is empty across every row when adaptive layout is enabled', () => {
    const table = buildTable(buildAdaptiveData(), true);

    // Stack Rank dropped — only ID and Title remain
    expect(table.Rows[0].Cells).toHaveLength(2);
    expect(table.Rows[1].Cells).toHaveLength(2);
    expect(table.Rows[2].Cells).toHaveLength(2);
  });

  test('never drops every column, even if all are empty', () => {
    const allEmpty = [
      {
        url: '',
        Source: 1,
        level: 1,
        fields: [
          { name: 'A', value: '' },
          { name: 'B', value: '' },
        ],
      },
    ];
    const table = buildTable(allEmpty, true);

    expect(table.Rows[0].Cells).toHaveLength(2);
  });

  test('assigns content-driven, non-empty percentage widths that sum to 100%', () => {
    const table = buildTable(buildAdaptiveData(), true);

    const headerWidths = table.Rows[0].Cells.map((c: any) => parseFloat(c.width));
    expect(headerWidths.every((w: number) => Number.isFinite(w) && w > 0)).toBe(true);

    const total = headerWidths.reduce((sum: number, w: number) => sum + w, 0);
    expect(total).toBeCloseTo(100, 0);

    // Title (long text) should end up wider than ID (short numeric values)
    const idWidth = headerWidths[0];
    const titleWidth = headerWidths[1];
    expect(titleWidth).toBeGreaterThan(idWidth);
  });
});

describe('JSONTable priority column ordering (enableAdaptiveLayout)', () => {
  const buildUnorderedData = () => [
    {
      url: '',
      Source: 1,
      level: 1,
      fields: [
        { name: 'Priority', value: 1 },
        { name: 'Work Item Type', value: 'Bug' },
        { name: 'State', value: 'Active' },
        { name: 'Title', value: 'Some bug to fix' },
        { name: 'ID', value: 1155 },
      ],
    },
  ];

  const buildTable = (data: any, enableAdaptiveLayout: boolean) =>
    new JSONTable(
      data,
      { isBold: true, IsItalic: false, IsUnderline: false, Size: 10, Uri: null as any, Font: 'Arial', InsertLineBreak: false, InsertSpace: false },
      { isBold: false, IsItalic: false, IsUnderline: false, Size: 10, Uri: null as any, Font: 'Arial', InsertLineBreak: false, InsertSpace: false },
      0,
      false,
      false,
      false,
      undefined,
      false,
      enableAdaptiveLayout
    ).getJSONTable() as any;

  const headerLabels = (table: any) => table.Rows[0].Cells.map((c: any) => c.Paragraphs[0].Runs[0].text);

  test('reorders ID/Title/Work Item Type/State first when adaptive layout is enabled', () => {
    const table = buildTable(buildUnorderedData(), true);
    expect(headerLabels(table)).toEqual(['ID', 'Title', 'Work Item Type', 'State', 'Priority']);
  });

  test('leaves column order untouched when adaptive layout is disabled (every other doc type)', () => {
    const table = buildTable(buildUnorderedData(), false);
    expect(headerLabels(table)).toEqual(['Priority', 'Work Item Type', 'State', 'Title', 'ID']);
  });

  test('is a no-op when columns are already in priority order', () => {
    const alreadyOrdered = [
      {
        url: '',
        Source: 1,
        level: 1,
        fields: [
          { name: 'ID', value: 1 },
          { name: 'Title', value: 'x' },
          { name: 'Priority', value: 1 },
        ],
      },
    ];
    const table = buildTable(alreadyOrdered, true);
    expect(headerLabels(table)).toEqual(['ID', 'Title', 'Priority']);
  });
});
