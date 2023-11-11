import RangeManager from './index';

describe('RangeManager', () => {
  beforeEach(() => {
    const people = [
      ["name", "age", "id", "is_active"],
      ["Alice", 28, 1, true],
      ["Bob", 35, 2, true],
      ["Charlie", 23, 3, false],
      ["David", 35, 4, true],
      ["Eva", 43, 5, true],
      ["Frank", 22, 6, true],
      ["Grace", 40, 7, false],
      ["Hank", 40, 8, true],
      ["Ivy", 29, 9, true],
      ["Jack", 37, 10, false]
    ];      

    global.SpreadsheetApp = {
      getActiveSpreadsheet: () => ({
        getSheetByName: () => ({
          getLastColumn: () => 3,
          getLastRow: () => 3,
          getRange: (row, column, numRows, numColumns) => ({
            getValues: () => {
              return people.slice(row - 1, row + numRows - 1).map((row) => row.slice(column - 1, column + numColumns - 1));
            },
            getValue: () => {
              return [people[row - 1][column - 1]];
            },
            setValues: (data) => {
              // Mock implementation for setValues
            },
            clearContent: () => {
              // Mock implementation for clearContent
            },
          }),
        }),
      }),
    };
  });

  afterEach(() => {
    delete global.SpreadsheetApp;
  });

  it('RangeManager is defined', () => {
    expect(RangeManager).toBeDefined();
  });

  it('fetch data using basic where clause', () => {
    const data = RangeManager.from('TestSheet').where({ name: 'Bob'}).fetch(['name', 'age']);
    expect(data).toEqual([{ name: 'Bob', age: 35 }]);
  });
});
