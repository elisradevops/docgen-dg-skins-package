import JSONParagraph from '../models/json/paragraph/JSONParagraph';

const baseStyles = {
  isBold: false,
  IsItalic: false,
  IsUnderline: false,
  Size: 12,
  Uri: null as string | null,
  Font: 'Arial',
  InsertLineBreak: false,
  InsertSpace: false,
};

describe('JSONParagraph label/value rendering', () => {
  test('splits a labeled field into an underlined label run, a non-underlined spacer run, and a value run', () => {
    const paragraph = new JSONParagraph(
      { name: 'Called Date', value: '07/12/2026 07:07:11' } as any,
      { ...baseStyles } as any,
      1,
      0,
    );
    const runs = paragraph.getJSONParagraph().runs;

    // Label run: underlined, ends with just the colon (no trailing space baked in)
    expect(runs[0].text).toBe('Called Date:');
    expect(runs[0].Underline).toBe(true);

    // Spacer run: a single space, NOT underlined — this is what stops the underline from
    // visually bleeding into the gap before the value starts.
    expect(runs[1].text).toBe(' ');
    expect(runs[1].Underline).toBe(false);

    // Value run: unaffected
    expect(runs[2].text).toBe('07/12/2026 07:07:11');
  });

  test('Title field renders as a single plain value run with no label/spacer', () => {
    const paragraph = new JSONParagraph(
      { name: 'Title', value: 'Something broke' } as any,
      { ...baseStyles } as any,
      1,
      0,
    );
    const runs = paragraph.getJSONParagraph().runs;

    expect(runs).toHaveLength(1);
    expect(runs[0].text).toBe('Something broke');
  });

  test('ID field renders as a single plain value run with no label/spacer', () => {
    const paragraph = new JSONParagraph(
      { name: 'ID', value: '1728' } as any,
      { ...baseStyles } as any,
      1,
      0,
    );
    const runs = paragraph.getJSONParagraph().runs;

    expect(runs).toHaveLength(1);
    expect(runs[0].text).toBe('1728');
  });
});
