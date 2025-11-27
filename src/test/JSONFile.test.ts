import JSONFile from '../models/json/file/JSONFile';
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

describe('JSONFile', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('removeFileExtension returns original name when there is no dot', () => {
    const data: any = {
      attachmentLink: 'http://server/file',
      attachmentFileName: 'file',
      attachmentType: 'asEmbedded',
      includeAttachmentContent: true,
    };

    const jsonFile = new JSONFile(data);
    const attachment = jsonFile.getFileAttachment as any;

    expect(attachment.type).toBe('File');
    expect(attachment.name).toBe('file');
  });

  test('isPictureUri handles invalid uri and logs error', () => {
    const data: any = {
      attachmentLink: undefined,
      attachmentFileName: 'file.docx',
      attachmentType: 'asEmbedded',
      includeAttachmentContent: true,
    };

    const jsonFile = new JSONFile(data);
    const attachment = jsonFile.getFileAttachment as any;

    expect(attachment.type).toBe('File');
    expect(attachment.name).toBe('file');
    expect(getMockLogger().error).toHaveBeenCalledWith(
      expect.stringContaining('error parsing file extension for')
    );
  });
});
