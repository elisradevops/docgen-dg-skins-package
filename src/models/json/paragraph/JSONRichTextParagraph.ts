import {
  Run,
  WIProperty,
  StyleOptions,
  RichNode,
  Paragraph,
  Table,
  TableRow,
  TableCell,
  List,
  ListItem,
} from '../wordJsonModels';

export default class JSONRichTextParagraph {
  paragraphStyles: StyleOptions;
  paragraphTemplate: any;
  wiId: number;
  constructor(
    field: WIProperty,
    paragraphStyles: StyleOptions,
    wiId: number,
    headingLevel: number,
    skipImageFormat = false
  ) {
    this.paragraphStyles = paragraphStyles;
    this.wiId = wiId;
    this.paragraphTemplate = this.buildDocFromNodes(field.richTextNodes);
  } //constructor

  getJSONRichTextParagraph(): any[] {
    return this.paragraphTemplate;
  } //getJSONTable

  /**
   * Parses a given RichNode and returns a corresponding Paragraph, Table, List,
   * an array of these types, or null if the node type is not recognized.
   *
   * @param node - The RichNode to be parsed.
   * @returns A Paragraph, Table, List, an array of these types, or null.
   *
   * The function handles different node types as follows:
   * - 'paragraph': Parses the node as a paragraph.
   * - 'table': Parses the node as a table.
   * - 'list': Parses the node as a list.
   * - 'text', 'image', 'break': Converts the node to a paragraph with a single run.
   * - 'other': Recursively parses child nodes and returns an array of results.
   * - Default: Returns null for unrecognized node types.
   */
  private parseNode(node: RichNode): Paragraph | Table | List | Array<Paragraph | Table | List> | null {
    switch (node.type) {
      case 'paragraph':
        return this.parseParagraphNode(node);
      case 'table':
        return this.parseTableNode(node);
      case 'list':
        return this.parseList(node);
      case 'text':
      case 'image':
      case 'break':
        return { type: 'paragraph', runs: [this.nodeToRun(node)] };
      case 'other':
        let childrenResults: Array<Paragraph | Table | List> = [];
        for (const child of node.children) {
          const childResult = this.parseNode(child);
          if (Array.isArray(childResult)) {
            childrenResults.push(...childResult);
          } else if (childResult) {
            childrenResults.push(childResult);
          }
        }
        return childrenResults.length > 0 ? childrenResults : null;
      default:
        return null;
    }
  }

  /**
   * Parses a paragraph node and converts it into a Paragraph object.
   *
   * @param node - The node to parse, expected to have a `children` property.
   * @returns A Paragraph object containing the parsed runs.
   */
  private parseParagraphNode(node: any): Paragraph {
    const runs: Run[] = this.gatherRuns(node.children);
    return { type: 'paragraph', runs };
  }

  /**
   * Gathers runs from the provided children nodes.
   *
   * This method processes an array of child nodes and converts them into an array of `Run` objects.
   * It handles different types of nodes such as text, image, break, paragraph, table, list, and other.
   *
   * - For `text`, `image`, and `break` nodes, it directly converts them to `Run` objects.
   * - For `paragraph` nodes, it recursively gathers runs from the child paragraphs.
   * - For `table` and `list` nodes, it adds a placeholder `Run` object indicating the embedded type.
   * - For `other` nodes, it recursively gathers runs from the child nodes.
   *
   * @param children - An array of child nodes to process.
   * @returns An array of `Run` objects representing the gathered runs.
   */
  private gatherRuns(children: any[]): Run[] {
    const runs: Run[] = [];

    for (const child of children || []) {
      if (child.type === 'text' || child.type === 'image' || child.type === 'break') {
        runs.push(this.nodeToRun(child));
      } else if (child.type === 'paragraph') {
        // Recursively gather runs from child paragraphs
        runs.push(...this.gatherRuns(child.children));
      } else if (child.type === 'table' || child.type === 'list') {
        // Not supported in this context, so we'll just add a placeholder
        runs.push({ type: 'other', value: `[Embedded ${child.type}]` });
      } else if (child.type === 'other') {
        // Recursively gather runs from inside <span> or <b>, etc.
        // TODO: Handle b and i, u tags and more relevant tags for text styling
        runs.push(...this.gatherRuns(child.children));
      }
    }
    return runs;
  }

  /**
   * Converts a node object to a Run object based on its type and text styling.
   *
   * @param node - The node object to be converted. It can be of type 'text', 'image', or 'break'.
   * @returns A Run object with the appropriate type and properties based on the input node.
   *
   * The Run object can have the following structure:
   * - For 'text' nodes: { type: 'text', value: string, textStyling: object }
   * - For 'image' nodes: { type: 'image', value: '', src: string }
   * - For 'break' nodes: { type: 'break' }
   * - For other types: { type: 'other', value: '' }
   *
   * The textStyling object includes:
   * - Bold: boolean
   * - Italic: boolean
   * - Underline: boolean
   * - Size: number
   * - Uri: string | null
   * - Font: string
   * - InsertLineBreak: boolean
   * - InsertSpace: boolean
   */
  private nodeToRun(node: any): Run {
    const textStyling = {
      Bold: node.textStyling?.Bold || this.paragraphStyles.isBold,
      Italic: node.textStyling?.Italic || this.paragraphStyles.IsItalic,
      Underline: node.textStyling?.Underline || this.paragraphStyles.IsUnderline,
      Size: this.paragraphStyles.Size || 12,
      Uri: this.paragraphStyles.Uri || null,
      Font: this.paragraphStyles.Font || 'Arial',
      InsertLineBreak: this.paragraphStyles.InsertLineBreak || false,
      InsertSpace: node.textStyling?.InsertSpace || this.paragraphStyles.InsertSpace,
    };

    if (node.type === 'text') {
      return { type: 'text', value: node.value, textStyling: textStyling };
    } else if (node.type === 'image') {
      return { type: 'image', value: '', src: node.src || '' };
    } else if (node.type === 'break') {
      return { type: 'break' };
    } else {
      return { type: 'other', value: '' };
    }
  }

  /**
   * Converts a given node object into a Table structure.
   *
   * @remarks
   * This function processes the children of the specified node and
   * transforms them into a table representation, adding default properties
   * such as heading level, page break insertion, and an empty rows array.
   *
   * @param node - The node object containing the structural details needed
   *               to create a table.
   * @returns A fully constructed Table object with the parsed node data.
   */
  private parseTableNode(node: any): Table {
    const table: Table = {
      type: 'table',
      headingLevel: 0,
      insertPageBreak: false,
      Rows: [],
    };

    this.parseTableChildren(node.children, table);
    return table;
  }

  /**
   * Parses child elements of a table structure and updates the provided Table instance accordingly.
   *
   * @param children - An array of child nodes to be evaluated for table content.
   * @param table - The Table object to which parsed rows and nested structures will be appended.
   *
   * @remarks
   * This method recursively traverses and identifies relevant table elements (e.g., <tr>, <thead>, <tbody>, etc.)
   * to generate the final row data for the given Table instance.
   */
  private parseTableChildren(children: any[], table: Table): void {
    for (const child of children || []) {
      if (child.type === 'other') {
        const tn = (child.tagName || '').toLowerCase();
        if (tn === 'tr') {
          // Parse a table row
          const row = this.parseTableRow(child);
          table.Rows.push(row);
        } else if (tn === 'thead' || tn === 'tbody' || tn === 'tfoot') {
          // Recurse into thead/tbody/tfoot
          this.parseTableChildren(child.children, table);
        } else {
          // Recurse into other nodes
          this.parseTableChildren(child.children, table);
        }
      } else if (child.type === 'table') {
        this.parseTableChildren(child.children, table);
      } else {
        //skip
        continue;
      }
    }
  }

  /**
   * Parses the given table row node and extracts its cells by iterating through its child elements.
   * Each child element that is recognized as a table cell (td or th) is converted into a corresponding cell object.
   *
   * @param trNode - An AST node representing the table row, containing possible table cell or header elements.
   * @returns A TableRow object containing an array of processed cells.
   */
  private parseTableRow(trNode: any): TableRow {
    const row: TableRow = { Cells: [] };

    // parse <td> or <th> inside this <tr>
    for (const child of trNode.children || []) {
      if (child.type === 'other' && (child.tagName === 'td' || child.tagName === 'th')) {
        row.Cells.push(this.parseTableCell(child));
      } else {
        // skip or recurse
      }
    }

    return row;
  }

  /**
   * Build one TableCell from <td> node
   */
  private parseTableCell(tdNode: any): TableCell {
    // We'll collect everything in one paragraph
    const cellParagraphRuns = this.gatherRunsFromAllParagraphs(tdNode.children);

    return {
      Paragraphs: [
        {
          Runs: cellParagraphRuns,
        },
      ],
    };
  }

  /**
   * Gather runs from any text/image/paragraph in <td>
   * (including nested <other> tags).
   */
  private gatherRunsFromAllParagraphs(children: any[]): Run[] {
    const runs: Run[] = [];
    for (const child of children || []) {
      if (child.type === 'text' || child.type === 'image') {
        runs.push(this.nodeToRun(child));
      } else if (child.type === 'paragraph') {
        runs.push(...this.gatherRuns(child.children));
      } else if (child.type === 'other' || child.type === 'table' || child.type === 'list') {
        // recursively gather
        runs.push(...this.gatherRunsFromAllParagraphs(child.children));
      }
    }
    return runs;
  }

  /**
   * Parses a given node to create a List object.
   *
   * @param node - The node to parse, expected to have properties `isOrdered` and `children`.
   * @returns A List object with the parsed data.
   */
  private parseList(node: any): List {
    const list: List = {
      type: 'list',
      isOrdered: node.isOrdered,
      listItems: [],
    };

    this.parseListChildren(node.children, list);
    return list;
  }

  /**
   * Parses the children of a list and adds them to the provided list object.
   *
   * @param children - An array of child nodes to be parsed.
   * @param list - The list object to which parsed items will be added.
   * @param level - The current depth level of the list (default is 0).
   */
  private parseListChildren(children: any[], list: List, level: number = 0): void {
    for (const child of children || []) {
      if (child.type === 'other') {
        const tn = (child.tagName || '').toLowerCase();
        if (tn === 'li') {
          // Parse a list item
          const item = this.parseListItem(child, level);
          list.listItems.push(item);
        } else {
          // Recurse into other nodes
          this.parseListChildren(child.children, list, level + 1);
        }
      } else {
        //skip
        continue;
      }
    }
  }

  /**
   * Parses a list item node and converts it into a `ListItem` object.
   *
   * @param liNode - The list item node to parse.
   * @param level - The nesting level of the list item.
   * @returns A `ListItem` object containing the parsed list item data.
   */
  private parseListItem(liNode: any, level: number): ListItem {
    const items = this.gatherRunsFromAllParagraphs(liNode.children);

    return {
      Runs: items,
      level: level,
    };
  }

  /**
   * Builds an array of document elements (Paragraph, Table, or List) from an array of RichNode objects.
   *
   * @param nodes - An array of RichNode objects to be parsed into document elements.
   * @returns An array of Paragraph, Table, or List elements.
   */
  private buildDocFromNodes(nodes: RichNode[]): Array<Paragraph | Table | List> {
    const results: Array<Paragraph | Table | List> = [];

    for (const node of nodes) {
      const items = this.parseNode(node);
      // parseNode might return a single Paragraph/Table or an array
      if (Array.isArray(items)) {
        results.push(...items);
      } else if (items) {
        results.push(items);
      }
    }

    return results;
  }
} //class
