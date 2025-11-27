import JSONTableRow from '../models/json/table/JSONTableRow';
import logger from '../services/logger';

jest.mock('../services/logger', () => ({
  __esModule: true,
  default: {
    debug: jest.fn(),
    info: jest.fn(),
    warn: jest.fn(),
    error: jest.fn(),
  },
}));

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

describe('JSONTableRow', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('throws error when WIData has no fields', () => {
    const data: any = {
      url: '',
      level: 0,
      Source: 1,
      // fields is intentionally missing to trigger the error path
    };

    expect(() => new JSONTableRow(data, { ...baseStyles } as any)).toThrow(
      'no fields to add for the table row'
    );
  });
});
