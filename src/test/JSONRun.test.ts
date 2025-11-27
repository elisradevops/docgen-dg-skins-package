import JSONRun from '../models/json/JSONRun';
import { defaultJsonRun } from '../models/json/default';
import logger from '../services/logger';

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
  isBold: true,
  IsItalic: false,
  IsUnderline: false,
  Size: 12,
  Uri: 'http://example.com',
  Font: 'Arial',
  InsertLineBreak: false,
  InsertSpace: true,
  InsertPageBreak: false as boolean | undefined,
};

describe('JSONRun', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('splits multi-line text into separate runs with correct line breaks', () => {
    const value = 'line1\nline2\nline3';

    const jsonRun = new JSONRun(value, { ...baseStyles });
    const runs = jsonRun.getRun();

    expect(runs).toHaveLength(3);

    expect(runs[0].InsertLineBreak).toBe(true);
    expect(runs[1].InsertLineBreak).toBe(true);
    expect(runs[2].InsertLineBreak).toBe(baseStyles.InsertLineBreak);

    runs.forEach((r) => {
      expect(r.Bold).toBe(baseStyles.isBold);
      expect(r.Italic).toBe(baseStyles.IsItalic);
      expect(r.Underline).toBe(baseStyles.IsUnderline);
      expect(r.Size).toBe(baseStyles.Size);
      expect(r.Font).toBe(baseStyles.Font);
      expect(r.InsertSpace).toBe(baseStyles.InsertSpace);
    });
  });

  test('uses striphtml replaceBr callback to convert <br> to newline', () => {
    const value = 'Hello<br>World';

    const jsonRun = new JSONRun(value, { ...baseStyles });
    const runs = jsonRun.getRun();

    expect(runs).toHaveLength(1);
    expect(runs[0].text).toContain('\n');
  });

  test('sets Uri to null when default run Uri is empty string in normal flow', () => {
    const prevUri = defaultJsonRun.Uri;
    defaultJsonRun.Uri = '' as any;

    const jsonRun = new JSONRun('plain text', { ...baseStyles });
    const runs = jsonRun.getRun();

    runs.forEach((r) => {
      expect(r.Uri).toBeNull();
    });

    defaultJsonRun.Uri = prevUri;
  });

  test('does not strip HTML when fieldtype is SuiteHeaderParagraphTitle', () => {
    const value = '<b>Bold Title</b>';

    const jsonRun = new JSONRun(value, { ...baseStyles }, 'SuiteHeaderParagraphTitle');
    const runs = jsonRun.getRun();

    expect(runs).toHaveLength(1);
    expect(runs[0].text).toContain('<b>Bold Title</b>');
  });

  test('falls back to default run and logs when value cannot be split', () => {
    const badValue: any = null;

    const prevUri = defaultJsonRun.Uri;
    defaultJsonRun.Uri = '' as any;

    const jsonRun = new JSONRun(badValue, { ...baseStyles });
    const runs = jsonRun.getRun();

    expect(getMockLogger().silly).toHaveBeenCalled();
    expect(runs).toHaveLength(1);
    expect(runs[0].text).toBe('');
    expect(runs[0].Bold).toBe(baseStyles.isBold);
    expect(runs[0].Italic).toBe(baseStyles.IsItalic);
    expect(runs[0].Underline).toBe(baseStyles.IsUnderline);
    expect(runs[0].Size).toBe(baseStyles.Size);
    expect(runs[0].Uri).toBeNull();

    defaultJsonRun.Uri = prevUri;
  });
});
