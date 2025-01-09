import { DocumentSkin, WIQueryResults, WIData, WIProperty, StyleOptions } from './models/json/wordJsonModels';
import logger from './services/logger';
import JSONTable from './models/json/table/JSONTable';
import JSONParagraph from './models/json/paragraph/JSONParagraph';
import JSONHeaderParagraph from './models/json/paragraph/JSONTestSuiteHeaderParagraph';
import JSONRichTextParagraph from './models/json/paragraph/JSONRichTextParagraph';
import { DescriptionandProcedureStyle } from './models/json/default';
import JSONFile from './models/json/file/JSONFile';

export default class Skins {
  SKIN_TYPE_TABLE = 'table';
  SKIN_TYPE_TRACE = 'trace-table';
  SKIN_TYPE_TABLE_STR = 'str-table';
  SKIN_TYPE_PARAGRAPH = 'paragraph';
  SKIN_TYPE_TEST_PLAN = 'test-plan';
  SKIN_TYPE_SYSTEM_OVERVIEW = 'system-overview';

  documentSkin: DocumentSkin = {
    templatePath: '',
    contentControls: [],
  };

  skinFormat: string = 'json';

  constructor(skinFormat: string = 'json', templatePath: string) {
    this.documentSkin.templatePath = templatePath;
    this.skinFormat = skinFormat;
  }

  async addNewContentToDocumentSkin(
    contentControlTitle: string,
    skinType: string,
    data: any,
    headerStyles: StyleOptions,
    styles: StyleOptions,
    headingLvl: number = 0,
    includeAttachments: boolean = true,
    insertPageBreak?: boolean,
    isFlattened: boolean = false
  ): Promise<any> {
    try {
      let populatedSkin;

      switch (skinType) {
        case this.SKIN_TYPE_TABLE:
          populatedSkin = this.generateQueryBasedTable(data, headerStyles, styles, headingLvl);
          if (populatedSkin === false) {
            logger.error(`Could not generate table for content control :${contentControlTitle}`);
            return false;
          }
          break;
        case this.SKIN_TYPE_TRACE:
          populatedSkin = this.generateTraceTable(data, headerStyles, styles, headingLvl);
          if (populatedSkin === false) {
            logger.error(`Could not generate table for content control :${contentControlTitle}`);
            return false;
          }
          break;
        case this.SKIN_TYPE_TABLE_STR:
          const insertPageBreakEnabled: boolean = insertPageBreak ?? false;
          populatedSkin = this.generateStrSkin(
            data,
            headerStyles,
            styles,
            headingLvl,
            insertPageBreakEnabled,
            contentControlTitle
          );
          if (populatedSkin === false) {
            logger.error(`Could not generate STR table for content control :${contentControlTitle}`);
            return false;
          }
          break;
        case this.SKIN_TYPE_PARAGRAPH:
          populatedSkin = this.generateQueryBasedParagraphs(data, styles, headingLvl);
          if (populatedSkin === false) {
            logger.error(`Could not generate paragraph for content control :${contentControlTitle}`);
            return false;
          }
          break;
        case this.SKIN_TYPE_TEST_PLAN:
          populatedSkin = this.generateTestBasedSkin(
            data,
            headerStyles,
            styles,
            headingLvl,
            includeAttachments,
            isFlattened
          );
          if (populatedSkin === false) {
            logger.error(`Could not generate test skin for content control :${contentControlTitle}`);
            return false;
          }
          break;
        case this.SKIN_TYPE_SYSTEM_OVERVIEW:
          populatedSkin = this.generateSystemOverviewSkin(data, styles);
          break;
        default:
          logger.error(`Unknown skinType : ${skinType} - not appended to document skin`);
          return false;
      }

      return populatedSkin;

      //  await this.validateAndAppendContentControl(
      //   contetnControlTitle,
      //   populatedSkin
      // );
    } catch (error) {
      logger.error(`Fatal error happened when generate skin for : ${contentControlTitle}`);
      logger.error(`Error Message: ${error.message}`);
      logger.error(`Error Stack: ${error.stack}`);
      return false;
    }
  } //addNewContentToDocumentSkin

  // iterates on content controls array:
  // if finds return true else creates content control and returns true
  // on error returns false
  async validateAndAppendContentControl(contetnControlTitle: string, populatedSkin: any): Promise<boolean> {
    try {
      let isContetnControlExists: boolean[] = await Promise.all(
        this.documentSkin.contentControls.map((contentControl, i) => {
          if (contentControl.title === contetnControlTitle) {
            logger.debug(`Content control - ${contetnControlTitle} found on index ${i}
          populating with: ${contetnControlTitle}`);
            this.documentSkin.contentControls[i].wordObjects = [
              ...this.documentSkin.contentControls[i].wordObjects,
              ...populatedSkin,
            ];
            return true;
          }
          return false;
        }) //.map
      ); //Promise.all

      if (!isContetnControlExists.includes(true)) {
        let contentControlToAppend = {
          title: contetnControlTitle,
          wordObjects: populatedSkin,
        };

        logger.debug(`Content control - ${contetnControlTitle} not found
           appending to document title:
           ${JSON.stringify(contetnControlTitle)}`);
        this.documentSkin.contentControls.unshift(contentControlToAppend);
      }
      logger.info(`Appended new content control document:  ${contetnControlTitle}`);
      return true;
    } catch (error) {
      logger.error(`Error validating and appending to content control`);
      console.log(error);
    }
  } //validateAndAppendContentControl

  //parses data to the class skin format and returns the parsed data
  //returns false if supplied an invalid type or error occurd during parsing
  generateQueryBasedTable(
    data: WIQueryResults,
    headerStyles: StyleOptions,
    styles: StyleOptions,
    headingLvl: number = 0
  ) {
    logger.debug(`Generating table as ${this.skinFormat}`);
    try {
      switch (this.skinFormat) {
        case 'json':
          let tableSkin = new JSONTable(data, headerStyles, styles, headingLvl);
          return [tableSkin.getJSONTable()];
        case 'html':
          logger.info(`Generating html table!`);
          break;
        default:
          return false;
      } //switch
    } catch (error: any) {
      logger.error(`Error occurred in generateQueryBasedTable: ${error.message}`);
      logger.error(`Error Stack: ${error.stack}`);
      return false;
    }
  } //generateTable

  private convertLevel = (headingLevel: number, level: number) => {
    if (level !== -1) {
      return level > 0 ? headingLevel + level : headingLevel ? headingLevel : 0;
    }
    return 0;
  };

  generateStrSkin(
    data: any,
    headerStyles: StyleOptions,
    styles: StyleOptions,
    headingLvl: number = 0,
    insertPageBreak: boolean,
    contentControlTitle: string
  ) {
    logger.debug(`Generating table as ${this.skinFormat}`);
    try {
      switch (this.skinFormat) {
        case 'json':
          if (
            contentControlTitle !== 'appendix-a-content-control' &&
            contentControlTitle !== 'appendix-b-content-control'
          ) {
            if (data.length > 0) {
              let tableSkin = new JSONTable(
                data,
                headerStyles,
                styles,
                headingLvl,
                undefined,
                insertPageBreak
              );
              return [tableSkin.getJSONTable()];
            } else {
              let emptyData = { name: 'Description', value: 'No relevant data' };
              let paragraphSkin = new JSONParagraph(emptyData, styles, 0, headingLvl);
              return [paragraphSkin.getJSONParagraph()];
            }
          } else {
            const strSkins = [];

            data.flat().forEach((skin) => {
              if (skin.type !== undefined && skin.type.includes('Header')) {
                // Title element, create a header paragraph
                const titleSkin = new JSONHeaderParagraph(
                  [skin.field],
                  {
                    isBold: true,
                    IsItalic: false,
                    IsUnderline: true,
                    Size: skin.type.includes('SubHeader') ? 12 : 14,
                    Uri: null,
                    Font: 'Arial',
                    InsertLineBreak: false,
                    InsertSpace: true,
                  },
                  0,
                  this.convertLevel(headingLvl, skin.level ?? -1)
                );
                strSkins.push(titleSkin.getJSONParagraph());
              } else if (skin.type === 'File') {
                const fileSkin = new JSONFile(skin);
                strSkins.push(fileSkin.getFileAttachment);
              } else {
                // Regular paragraph element
                const paragraph = new JSONParagraph(
                  skin.field,
                  styles, // Assuming it's defined elsewhere
                  0,
                  0
                );
                strSkins.push(paragraph.getJSONParagraph());
              }
            });
            return strSkins;
          }
        case 'html':
          logger.info(`Generating html table!`);
          break;
        default:
          return false;
      } //switch
    } catch (error: any) {
      logger.error(`Error occurred in generateStrBasedTable: ${error.message}`);
      logger.error(`Error Stack: ${error.stack}`);
      return false;
    }
  } //generateTable

  generateSystemOverviewSkin(systemOverviewAdaptedData: any[], styles: StyleOptions) {
    let headerStyle = {
      isBold: true,
      IsItalic: false,
      IsUnderline: false,
      Size: 14,
      Uri: null,
      Font: 'Arial',
      InsertLineBreak: false,
      InsertSpace: true,
    };
    let testSkins: any[] = [];
    for (const element of systemOverviewAdaptedData) {
      let wiHeaderFields = [element.fields[0], element.fields[1]];

      let wiSkin = new JSONHeaderParagraph(wiHeaderFields, headerStyle, element.id || 0, element.level);
      testSkins.push(wiSkin.getJSONParagraph());

      //TODO: remove this hardcoding
      let wiDescTitleSkin = new JSONParagraph(
        { name: 'Title', value: 'WI Description' },
        DescriptionandProcedureStyle,
        element.id || 0,
        0
      );
      testSkins.push(wiDescTitleSkin.getJSONParagraph());
      let wiDescriptionParagraph = new JSONRichTextParagraph(element.fields[2], styles, element.id || 0, 0);
      let richTextDecsSkin: any[] = wiDescriptionParagraph.getJSONRichTextParagraph();
      testSkins.push(...richTextDecsSkin);
    }
    return testSkins;
  }

  generateQueryBasedParagraphs(data: any, styles: StyleOptions, headingLvl: number = 0) {
    logger.debug(`Generating paragraph as ${this.skinFormat}`);

    switch (this.skinFormat) {
      case 'json':
        let paragraphs: any[] = [];
        data.forEach((wi: WIData) => {
          wi.fields.forEach((field: WIProperty) => {
            //Description and Test Description are handled in richText handler
            if (field.name === 'Description: ' || field.name === 'Test Description:') {
              let paragraphSkin = new JSONRichTextParagraph(field, styles, wi.Source, headingLvl + wi.level);
              paragraphs.push(paragraphSkin.getJSONRichTextParagraph());
            }

            //making sure ID is not printed
            //excluding Description and Test Description are handled in richText handler
            if (field.name !== 'ID' && field.name !== 'Description: ' && field.name !== 'Test Description:') {
              let paragraphSkin = new JSONParagraph(field, styles, wi.Source, headingLvl + wi.level);
              paragraphs.push(paragraphSkin.getJSONParagraph());
            }
          });
        });
        return paragraphs;
      case 'html':
        logger.info(`Generating html paragraphs!`);
        break;
      default:
        return false;
    } //switch
  } //generateParagraph

  generateTraceTable(data: any, headerStyles: StyleOptions, styles: StyleOptions, headingLvl: number = 0) {
    logger.debug(`Generating table as ${this.skinFormat}`);
    try {
      switch (this.skinFormat) {
        case 'json':
          let traceSkin: any[] = [];
          const { title, adoptedData, errorMessage } = data;
          let traceTitle = new JSONHeaderParagraph(title.fields, styles, undefined, 2);
          traceSkin.push(traceTitle.getJSONParagraph());

          if (adoptedData?.length > 0) {
            let tableSkin = new JSONTable(adoptedData, headerStyles, styles, headingLvl);
            traceSkin.push(tableSkin.getJSONTable());
          } else if (errorMessage !== null) {
            let errorSkin = new JSONParagraph({ name: 'Description', value: errorMessage }, styles, 0, 0);
            traceSkin.push(errorSkin.getJSONParagraph());
          }
          return traceSkin;
        case 'html':
          logger.info(`Generating html table!`);
          break;
        default:
          return false;
      } //switch
    } catch (error: any) {
      logger.error(`Error occurred in generateTraceTable: ${error.message}`);
      logger.error(`Error Stack: ${error.stack}`);
      return false;
    }
  } //generateTraceTable

  generateTestBasedSkin(
    data: any,
    headerStyles: StyleOptions,
    styles: StyleOptions,
    headingLvl: number = 0,
    includeAttachments: boolean = true,
    isFlattened = false
  ) {
    logger.debug(`Generating testSkin as ${this.skinFormat}`);
    switch (this.skinFormat) {
      case 'json':
        let testSkin: any[] = [];
        //create suite Header pragraph
        data.forEach((testSuite: any) => {
          let SuiteStyles = {
            isBold: true,
            IsItalic: false,
            IsUnderline: false,
            Size: 16,
            Uri: null,
            Font: 'Arial',
            InsertLineBreak: false,
            InsertSpace: true,
          };
          let testSuiteParagraphSkin = new JSONHeaderParagraph(
            testSuite.suiteSkinData.fields,
            SuiteStyles,
            testSuite.suiteSkinData.id || 0,
            testSuite.suiteSkinData.level
              ? headingLvl + testSuite.suiteSkinData.level
              : headingLvl
              ? headingLvl
              : 1
          );
          testSkin.push(testSuiteParagraphSkin.getJSONParagraph());

          testSuite.testCases.forEach((testcase) => {
            //create testcase Header paragraph
            let testCaseHeaderFields = [
              testcase.testCaseHeaderSkinData.fields[0],
              testcase.testCaseHeaderSkinData.fields[1],
            ];
            let TestCaseStyles = {
              isBold: true,
              IsItalic: false,
              IsUnderline: false,
              Size: 14,
              Uri: null,
              Font: 'Arial',
              InsertLineBreak: false,
              InsertSpace: true,
            };
            let testCaseParagraphSkin = new JSONHeaderParagraph(
              testCaseHeaderFields,
              TestCaseStyles,
              testcase.id || 0,
              testcase.testCaseHeaderSkinData.level
                ? headingLvl + testcase.testCaseHeaderSkinData.level
                : headingLvl
                ? headingLvl
                : 1
            );
            testSkin.push(testCaseParagraphSkin.getJSONParagraph());

            let testDescriptionTitleParagraph = new JSONParagraph(
              { name: 'Title', value: 'Test Description:' },
              DescriptionandProcedureStyle,
              testcase.id || 0,
              0
            );
            testSkin.push(testDescriptionTitleParagraph.getJSONParagraph());

            let testDescriptionParagraph = new JSONRichTextParagraph(
              testcase.testCaseHeaderSkinData.fields[2],
              { ...styles, Size: 12, Font: 'Arial' },
              testcase.id || 0,
              0
            );
            let richTextSkin: any[] = testDescriptionParagraph.getJSONRichTextParagraph();
            testSkin = [...testSkin, ...richTextSkin];

            if (testcase.testCaseRequirements) {
              if (testcase.testCaseRequirements.length > 0) {
                let testDescriptionTitleParagraph = new JSONParagraph(
                  { name: 'Title', value: 'Covered Requirements:' },
                  DescriptionandProcedureStyle,
                  testcase.id || 0,
                  0
                );
                testSkin.push(testDescriptionTitleParagraph.getJSONParagraph());
                //create test steps table
                let tableSkin = new JSONTable(
                  testcase.testCaseRequirements,
                  headerStyles,
                  styles,
                  headingLvl
                );

                let populatedTableSkin = tableSkin.getJSONTable();

                testSkin.push(populatedTableSkin);
              }
            }

            if (testcase.testCaseBugs) {
              if (testcase.testCaseBugs.length > 0) {
                let testDescriptionTitleParagraph = new JSONParagraph(
                  { name: 'Title', value: 'Linked Bugs:' },
                  DescriptionandProcedureStyle,
                  testcase.id || 0,
                  0
                );
                testSkin.push(testDescriptionTitleParagraph.getJSONParagraph());
                //create test steps table
                let tableSkin = new JSONTable(testcase.testCaseBugs, headerStyles, styles, headingLvl);

                let populatedTableSkin = tableSkin.getJSONTable();

                testSkin.push(populatedTableSkin);
              }
            }

            try {
              if (testcase.testCaseStepsSkinData.length > 0) {
                let testProcedureTitleParagraph = new JSONParagraph(
                  { name: 'Title', value: 'Test Procedure:' },
                  DescriptionandProcedureStyle,
                  testcase.id || 0,
                  0
                );
                testSkin.push(testProcedureTitleParagraph.getJSONParagraph());
                //create test steps table
                let tableSkin = new JSONTable(
                  testcase.testCaseStepsSkinData,
                  headerStyles,
                  { ...styles, Font: 'Arial' },
                  headingLvl,
                  undefined,
                  undefined,
                  isFlattened
                );
                let populatedTableSkin = tableSkin.getJSONTable();
                testSkin.push(populatedTableSkin);
              }
            } catch (error) {
              logger.warn(
                `For suite id : ${testSuite.suiteSkinData.fields[0].value} , the testCaseStepsSkinData is not defined for ${testcase.testCaseHeaderSkinData.fields[0].value} `
              );
            }

            //attachments table
            try {
              if (testcase.testCaseAttachments) {
                if (testcase.testCaseAttachments.length > 0 && includeAttachments) {
                  let testDescriptionTitleParagraph = new JSONParagraph(
                    { name: 'Title', value: 'Test Case Attachments:' },
                    styles,
                    testcase.id || 0,
                    0
                  );
                  testSkin.push(testDescriptionTitleParagraph.getJSONParagraph());
                  //create test steps table

                  let tableSkin = new JSONTable(
                    testcase.testCaseAttachments,
                    headerStyles,
                    { ...styles, Font: 'Arial' },
                    headingLvl,
                    true
                  );
                  let populatedTableSkin = tableSkin.getJSONTable();
                  testSkin.push(populatedTableSkin);
                }
              }
            } catch (error: any) {
              logger.error(`Error occurred when building attachments ${error.message}}`);
              logger.error(`Stack: ${error.stack}}`);
            }
          });
        });
        return testSkin;
        break;
      case 'html':
        logger.info(`Generating html test data!`);
        break;
      default:
        return false;
    } //switch
  } //generateTestBasedSkin

  getDocumentSkin(): DocumentSkin {
    return this.documentSkin;
  } //getDocumentSkin
}
