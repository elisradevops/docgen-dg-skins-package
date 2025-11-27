import JSONTableCell from '../models/json/table/JSONTableCell';
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

const getMockLogger = () =>
  logger as unknown as {
    debug: jest.Mock;
    info: jest.Mock;
    warn: jest.Mock;
    error: jest.Mock;
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

describe('JSONTableCell', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('creates picture attachment when value contains picture link and respects isFlattened', () => {
    const shading = { color: 'auto', fill: 'FF0000', themeFillShade: 'BF' } as any;

    const data: any = {
      name: 'AttachmentCell',
      value: [
        {
          attachmentLink: 'http://server/image.png',
          tableCellAttachmentLink: 'http://server/image-render.png',
          attachmentFileName: 'image.png',
        },
      ],
      attachmentType: 'asEmbedded',
      includeAttachmentContent: true,
      width: '50%',
    };

    const cell = new JSONTableCell(data, { ...baseStyles }, false, shading, true);
    const json = cell.getJsonCell() as any;

    expect(json.shading).toEqual(shading);
    expect(json.attachments).toHaveLength(1);
    expect(json.attachments[0]).toMatchObject({
      type: 'Picture',
      path: 'http://server/image-render.png',
      name: 'image',
      isFlattened: true,
    });
  });

  test('creates file attachment with link flags for non-picture attachment', () => {
    const data: any = {
      name: 'AttachmentCell',
      value: [
        {
          attachmentLink: 'http://server/doc1.pdf',
          attachmentFileName: 'doc1.pdf',
        },
      ],
      attachmentType: 'asLink',
      includeAttachmentContent: false,
      width: '50%',
    };

    const cell = new JSONTableCell(data, { ...baseStyles }, false, undefined, true);
    const json = cell.getJsonCell() as any;

    expect(json.attachments).toHaveLength(1);
    expect(json.attachments[0]).toMatchObject({
      type: 'File',
      path: 'http://server/doc1.pdf',
      name: 'doc1',
      isLinkedFile: true,
      includeAttachmentContent: false,
      isFlattened: true,
    });
  });

  test('creates HTML cell when value contains HTML markup', () => {
    const shading = { color: 'auto', fill: '00FF00' } as any;

    const data: any = {
      name: 'HtmlCell',
      value: '<p>Hello</p>',
      width: '30%',
    };

    const cell = new JSONTableCell(data, { ...baseStyles }, false, shading, false);
    const json = cell.getJsonCell() as any;

    expect(json.Html).toMatchObject({
      type: 'html',
      Html: '<p>Hello</p>',
      font: 'Arial',
      fontSize: 10,
    });
    expect(json.shading).toEqual(shading);
    expect(json.attachments).toHaveLength(0);
  });

  test('creates text runs when value is simple text', () => {
    const data: any = {
      name: 'TextCell',
      value: 'Plain text',
      width: '20%',
      url: 'http://example.com/field',
    };

    const cell = new JSONTableCell(data, { ...baseStyles }, false, undefined, false);
    const json = cell.getJsonCell() as any;

    expect(json.Html).toBeUndefined();
    expect(json.attachments).toHaveLength(0);
    expect(json.Paragraphs[0].Runs.length).toBeGreaterThan(0);
    const text = json.Paragraphs[0].Runs.map((r: any) => r.text).join('');
    expect(text).toContain('Plain text');
  });

  test('logs error and falls back to error run when value is invalid non-array object', () => {
    const data: any = {
      name: 'BadCell',
      value: { not: 'an array' },
      width: '10%',
    };

    const cell = new JSONTableCell(data, { ...baseStyles }, false, undefined, false);
    const json = cell.getJsonCell() as any;

    expect(getMockLogger().error).toHaveBeenCalled();
    const text = json.Paragraphs[0].Runs.map((r: any) => r.text).join('');
    expect(text).toContain('Docgen Error: Invalid data value');
  });
});
