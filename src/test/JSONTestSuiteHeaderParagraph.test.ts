import JSONTestSuiteHeaderParagraph from '../models/json/paragraph/JSONTestSuiteHeaderParagraph';
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
  isBold: true,
  IsItalic: false,
  IsUnderline: false,
  Size: 14,
  Uri: null as string | null,
  Font: 'Arial',
  InsertLineBreak: false,
  InsertSpace: true,
};

describe('JSONTestSuiteHeaderParagraph', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('creates runs for non-title/id/description fields and sets heading level', () => {
    const fields: any[] = [{ name: 'Custom', value: 'Value', url: 'http://example.com' }];

    const paragraph = new JSONTestSuiteHeaderParagraph(fields, { ...baseStyles }, 1, 2);
    const json = paragraph.getJSONParagraph() as any;

    expect(json.type).toBe('paragraph');
    expect(json.headingLevel).toBe(2);
    expect(json.runs.length).toBeGreaterThan(0);
  });
});
