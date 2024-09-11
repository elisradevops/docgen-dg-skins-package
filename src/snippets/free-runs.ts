import Skins from '../index';

let main = async () => {
  let headerStyles = {
    isBold: true,
    IsItalic: false,
    IsUnderline: false,
    Size: 12,
    Uri: null,
    Font: 'Arial',
    InsertLineBreak: false,
    InsertSpace: true,
  };

  let styles = {
    isBold: false,
    IsItalic: false,
    IsUnderline: false,
    Size: 12,
    Uri: null,
    Font: 'Arial',
    InsertLineBreak: false,
    InsertSpace: true,
  };

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
};

main();

// let json = {
//   workItems: [
//     {
//       //TEST PLAN ROOT
//       url: "test-plan-url",
//       fields: [
//         { value: " test-plan-title", name: "title" },
//         { value: "test-plan-summary", name: "test-plan-summary" },
//       ],
//       parent: null,
//       id: "test-plan-id",
//     },
//     {
//       //TEST SUITE
//       url: "test-suite-url",
//       fields: [
//         { value: " test-suite-title", name: "title" },
//         { value: "test-suite-summary", name: "test-suite-summary" },
//         { value: "User Story", name: "Work Item Type" },
//       ],
//       parent: "suite-id",
//       id: "test-plan-id",
//     },
//   ],
// };
