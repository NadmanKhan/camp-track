import {
  problemIndexToLetter,
  problemLetterToIndex,
} from '../src/main';

describe('problemIndexToLetter', () => {
  it('should return the corresponding letter for index < 26', () => {
    expect(problemIndexToLetter(0)).toBe('A');
    expect(problemIndexToLetter(1)).toBe('B');
    expect(problemIndexToLetter(25)).toBe('Z');
  });

  it('should return the correct 2-letter string for index >= 26 and < 702', () => {
    expect(problemIndexToLetter(26)).toBe('AA');
    expect(problemIndexToLetter(27)).toBe('AB');
    expect(problemIndexToLetter(51)).toBe('AZ');
    expect(problemIndexToLetter(52)).toBe('BA');
    expect(problemIndexToLetter(53)).toBe('BB');
    expect(problemIndexToLetter(77)).toBe('BZ');
    expect(problemIndexToLetter(78)).toBe('CA');
    expect(problemIndexToLetter(79)).toBe('CB');
    expect(problemIndexToLetter(699)).toBe('ZX');
    expect(problemIndexToLetter(700)).toBe('ZY');
    expect(problemIndexToLetter(701)).toBe('ZZ');
  });

  it('should throw an error for non-integer index', () => {
    expect(() => problemIndexToLetter(1.5)).toThrow();
  });

  it('should throw an error for negative index', () => {
    expect(() => problemIndexToLetter(-1)).toThrow();
  });

  it('should throw an error for index >= 702', () => {
    expect(() => problemIndexToLetter(702)).toThrow();
  });
});

describe('problemLetterToIndex', () => {
  it('should return the corresponding index for letter < "AA"', () => {
    expect(problemLetterToIndex('A')).toBe(0);
    expect(problemLetterToIndex('B')).toBe(1);
    expect(problemLetterToIndex('Z')).toBe(25);
  });

  it('should return the corresponding index for letter >= "AA" and < "ZZ"', () => {
    expect(problemLetterToIndex('AA')).toBe(26);
    expect(problemLetterToIndex('AB')).toBe(27);
    expect(problemLetterToIndex('AZ')).toBe(51);
    expect(problemLetterToIndex('BA')).toBe(52);
    expect(problemLetterToIndex('BB')).toBe(53);
    expect(problemLetterToIndex('BZ')).toBe(77);
    expect(problemLetterToIndex('CA')).toBe(78);
    expect(problemLetterToIndex('CB')).toBe(79);
    expect(problemLetterToIndex('ZX')).toBe(699);
    expect(problemLetterToIndex('ZY')).toBe(700);
    expect(problemLetterToIndex('ZZ')).toBe(701);
  });

  it('should throw an error for empty letter', () => {
    expect(() => problemLetterToIndex('')).toThrow();
  });

  it('should throw an error for letter containing non-letters', () => {
    expect(() => problemLetterToIndex('A1')).toThrow();
    expect(() => problemLetterToIndex('.A')).toThrow();
    expect(() => problemLetterToIndex('B-A')).toThrow();
    expect(() => problemLetterToIndex('A!')).toThrow();
  });
});
