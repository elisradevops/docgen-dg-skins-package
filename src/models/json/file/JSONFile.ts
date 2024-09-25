import { extname } from 'path';
import { Attachment } from '../wordJsonModels';
import logger from '../../../services/logger';

export default class JSONFile {
  private _fileAttachment: Attachment;

  constructor(fileAttachmentData: any) {
    this._fileAttachment = this.generateFileAttachment(fileAttachmentData);
  }

  public get getFileAttachment(): Attachment {
    return this._fileAttachment;
  }

  private generateFileAttachment(fileAttachmentData: any): Attachment {
    if (this.isPictureUri(fileAttachmentData.attachmentLink)) {
      return {
        type: 'Picture',
        path: fileAttachmentData.attachmentLink,
        name: this.removeFileExtension(fileAttachmentData.attachmentFileName),
      };
    } else {
      return {
        type: 'File',
        path: fileAttachmentData.attachmentLink,
        name: this.removeFileExtension(fileAttachmentData.attachmentFileName),
        isLinkedFile: fileAttachmentData.attachmentType === 'asLink',
      };
    }
  }

  private removeFileExtension(filePath) {
    const lastDotIndex = filePath.lastIndexOf('.');
    if (lastDotIndex > 0) {
      return filePath.slice(0, lastDotIndex);
    }
    return filePath; // If there's no extension, return the original file path
  }

  private isPictureUri(uri: string) {
    try {
      let fileExt = extname(uri);
      logger.debug(`attachment extension: ${fileExt}`);
      if (
        fileExt.toUpperCase() == '.PNG' ||
        fileExt.toUpperCase() == '.JPG' ||
        fileExt.toUpperCase() == '.JPEG' ||
        fileExt.toUpperCase() == '.GIF'
      ) {
        return true;
      } else {
        return false;
      }
    } catch (e) {
      logger.error(`error parsing file extension for : ${uri}`);
      return false;
    }
  }
}
