import { writeFileSync } from 'fs';
import Skins from '../index';
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

describe('Generate json skins from queries - tests', () => {
  test('generate table skin - flat query', async () => {
    let skins = new Skins('json', 'c:\test.dotx');
    const QueryData = require('../../samples/mock-data/queries/flat-query-data.json');
    await skins.addNewContentToDocumentSkin(
      'test',
      skins.SKIN_TYPE_TABLE,
      QueryData,
      headerStyles,
      styles,
      4
    );
    const documentSkin: any = skins.getDocumentSkin();
    const SnapShot = require('../../samples/snapshots/queries/flat-query-table-snapshot.json');
    expect(documentSkin).toMatchObject(SnapShot);
  });
  test('generate table skin - tree query', async () => {
    let skins = new Skins('json', 'c:\test.dotx');
    const QueryData = require('../../samples/mock-data/queries/tree-query-data.json');
    await skins.addNewContentToDocumentSkin(
      'test',
      skins.SKIN_TYPE_TABLE,
      QueryData,
      headerStyles,
      styles,
      4
    );
    const documentSkin: any = skins.getDocumentSkin();
    const SnapShot = require('../../samples/snapshots/queries/tree-query-table-snapshot.json');
    expect(documentSkin).toMatchObject(SnapShot);
  });
  test('generate paragraph skin - flat query', async () => {
    let skins = new Skins('json', 'c:\test.dotx');
    const QueryData = require('../../samples/mock-data/queries/flat-query-data.json');
    await skins.addNewContentToDocumentSkin(
      'test',
      skins.SKIN_TYPE_PARAGRAPH,
      QueryData,
      headerStyles,
      styles,
      4
    );
    const documentSkin: any = skins.getDocumentSkin();
    // writeFileSync(
    //   "flat-query-paragraph-snapshot.json",
    //   JSON.stringify(documentSkin)
    // );
    const SnapShot = require('../../samples/snapshots/queries/flat-query-paragraph-snapshot.json');
    expect(documentSkin).toMatchObject(SnapShot);
  });
  test('generate paragraph skin - tree query', async () => {
    let skins = new Skins('json', 'c:\test.dotx');
    const QueryData = require('../../samples/mock-data/queries/tree-query-data.json');
    await skins.addNewContentToDocumentSkin(
      'test',
      skins.SKIN_TYPE_PARAGRAPH,
      QueryData,
      headerStyles,
      styles,
      4
    );
    const documentSkin: any = skins.getDocumentSkin();
    // writeFileSync("tree-query-data.json", JSON.stringify(documentSkin));
    const SnapShot = require('../../samples/snapshots/queries/tree-query-paragraph-snapshot.json');
    expect(documentSkin).toMatchObject(SnapShot);
  });
});
describe('Generate json skins from testData - tests', () => {
  test('generate std skin - testData', async () => {
    let skins = new Skins('json', 'c:\\\\test\\\\test.dotx');
    const QueryData = require('../../samples/mock-data/test-data/testContentData.json');
    await skins.addNewContentToDocumentSkin(
      'system-capabilities',
      skins.SKIN_TYPE_TEST_PLAN,
      QueryData,
      headerStyles,
      styles,
      4
    );
    const documentSkin: any = skins.getDocumentSkin();

    const SnapShot = require('../../samples/snapshots/tests/flat-test-std-snapshot.json');
    expect(documentSkin).toMatchObject(SnapShot);
  });
  test('generate std skin - complex testData', async () => {
    let skins = new Skins('json', 'C:\\\\docgen\\\\documents\\\\181020205911\\\\test.dotx');
    const QueryData = require('../../samples/mock-data/test-data/complex-test-data-with-attachments.json');
    await skins.addNewContentToDocumentSkin(
      'system-capabilities',
      skins.SKIN_TYPE_TEST_PLAN,
      QueryData,
      headerStyles,
      styles,
      4
    );
    const documentSkin: any = skins.getDocumentSkin();
    const SnapShot = require('../../samples/snapshots/tests/complex-test-std-snapshot.json');
    expect(documentSkin).toMatchObject(SnapShot);
  });
  test('generate trace table skin', async () => {
    let skins = new Skins('json', 'c:\\test\\test.dotx');
    const QueryData = require('../../samples/mock-data/test-data/test-trace-table-data.json');

    await skins.addNewContentToDocumentSkin(
      'system-capabilities',
      skins.SKIN_TYPE_TABLE,
      QueryData,
      headerStyles,
      styles,
      4
    );
    const documentSkin: any = skins.getDocumentSkin();
    const SnapShot = require('../../samples/snapshots/tests/test-trace-table-snapshot.json');
    expect(documentSkin).toMatchObject(SnapShot);
  });
});
describe('Generate json skins svd - tests', () => {
  test('generate changes between releases skin', async () => {
    let skins = new Skins('json', 'c:\\test\\test.dotx');
    const adoptedChangesData = require('../../samples/mock-data/changes/adoptedChanges.json');

    for (let i = 0; i < adoptedChangesData.length; i++) {
      let artifactChangesData = adoptedChangesData[i];
      await skins.addNewContentToDocumentSkin(
        'change-description-content-control',
        skins.SKIN_TYPE_PARAGRAPH,
        artifactChangesData.artifact,
        headerStyles,
        styles,
        4
      );

      await skins.addNewContentToDocumentSkin(
        'change-description-content-control',
        skins.SKIN_TYPE_TABLE,
        artifactChangesData.artifactChanges,
        headerStyles,
        styles,
        4
      );
    }

    const documentSkin: any = skins.getDocumentSkin();
    const SnapShot = require('../../samples/snapshots/changes/changes-table-release-snapshot.json');
    expect(documentSkin).toMatchObject(SnapShot);
  });
});
describe('Common functions - tests', () => {
  test('generate richText skin', async () => {
    let skins = new Skins('json', 'c:\\\\test\\\\test.dotx');
    const QueryData = require('../../samples/mock-data/test-data/testContentDataWithRichText.json');
    await skins.addNewContentToDocumentSkin(
      'system-capabilities',
      skins.SKIN_TYPE_TEST_PLAN,
      QueryData,
      headerStyles,
      styles,
      4
    );
    const documentSkin: any = skins.getDocumentSkin();
    const SnapShot = require('../../samples/snapshots/common/rich-text-test-std.json');
    expect(documentSkin).toMatchObject(SnapShot);
  });
});
