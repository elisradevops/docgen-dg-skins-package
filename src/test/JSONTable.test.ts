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
