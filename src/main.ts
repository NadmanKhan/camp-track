
// Webpack GasPlugin will turn this into a dummy top level function for GAS:
global.main = () => { };

// ==================== Constants ====================

const BASE_URL = 'https://vjudge.net';
const CONTESTS_URL = `${BASE_URL}/contest`;

const SHEET_CONFIG: SheetConfig = {
  participants: {
    sheetName: 'Participants',
    first: { row: 2, column: 2 },
    schema: ['email', 'name', 'id', 'vjudgeUsernames'],
  },
  tasks: {
    sheetName: 'Tasks',
    first: { row: 2, column: 1 },
    schema: ['contest.id', 'solveTarget.total', 'solveTarget.problems'],
  },
  progress: {
    sheetName: 'Progress',
  },
} as const;

const COLOR_VALUES = {
  blue: '#0000ff',
  cornflower_blue: '#4a86e8',
  cyan: '#00ffff',
  green: '#00ff00',
  yellow: '#ffff00',
  orange: '#ff9900',
  red: '#ff0000',
  red_berry: '#980000',
  purple: '#9900ff',
  magenta: '#ff00ff',

  light_blue_1: '#6fa8dc',
  light_cornflower_blue_1: '#6d9eeb',
  light_cyan_1: '#76a5af',
  light_green_1: '#93c47d',
  light_yellow_1: '#ffd966',
  light_orange_1: '#f6b26b',
  light_red_1: '#e06666',
  light_red_berry_1: '#cc4125',
  light_purple_1: '#7e6bc4',
  light_magenta_1: '#c27ba0',

  light_blue_2: '#9fc5e8',
  light_cornflower_blue_2: '#a4c2f4',
  light_cyan_2: '#a2c4c9',
  light_green_2: '#b6d7a8',
  light_yellow_2: '#ffe599',
  light_orange_2: '#f9cb9c',
  light_red_2: '#ea9999',
  light_red_berry_2: '#dd7e6b',
  light_purple_2: '#b4a7d6',
  light_magenta_2: '#d5a6bd',

  light_blue_3: '#cfe2f3',
  light_cornflower_blue_3: '#c9daf8',
  light_cyan_3: '#d0e0e3',
  light_green_3: '#d9ead3',
  light_yellow_3: '#fff2cc',
  light_orange_3: '#fce5cd',
  light_red_3: '#f4cccc',
  light_red_berry_3: '#e6b8af',
  light_purple_3: '#d9d2e9',
  light_magenta_3: '#ead1dc',

  white: '#ffffff',
  light_grey_3: '#f3f3f3',
  light_grey_2: '#efefef',
  light_grey_1: '#d9d9d9',
  grey: '#cccccc',
  dark_grey_1: '#b7b7b7',
  dark_grey_2: '#999999',
  dark_grey_3: '#666666',
  black: '#000000',
} as const;

const COLUMN_WIDTHS = {
  xs: 20,
  sm: 50,
  md: 90,
  lg: 140,
  xl: 200
} as const;


// ==================== Main function ====================

(function main() {

  const participants = readParticipants();

  const tasks = readTasks();

  const progressMap = computeProgressMap(participants, tasks);

  const rankedParticipants = computeRanking(participants, progressMap);

  const taskAggreate = aggregateTasks(tasks);

  writeProgressSheet(rankedParticipants, taskAggreate);
})();


// ==================== Read from sheet ====================

export function readParticipants(): Participant[] {
  const participants: Participant[] = [];

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEET_CONFIG.participants.sheetName);

  const { first, schema } = SHEET_CONFIG.participants;

  for (let row = first.row; ; ++row) {
    if (sheet.getRange(row, first.column).getValue() === '') {
      break;
    }

    const participant: Participant = {
      id: '',
      name: '',
      email: '',
      vjudgeUsernames: [],
    };

    for (let column = first.column; column < first.column + schema.length; ++column) {
      const value = sheet.getRange(row, column).getValue();
      switch (schema[column - first.column]) {
        case 'id':
        case 'name':
        case 'email':
          if (typeof value === 'string') {
            participant[schema[column - first.column]] = value;
          } else {
            participant[schema[column - first.column]] = value.toString();
          }
          break;
        case 'vjudgeUsernames':
          participant.vjudgeUsernames = (value as string).split(',')
            .map(x => x.trim())
            .filter(x => x.length > 0)
            .sort();
          break;
        default:
          throw new Error(`Unknown schema: ${schema[column - first.column]}`);
      }
    }

    Logger.log(participant);
    participants.push(participant);
  }

  return participants;
}

export function readTasks(): Task[] {
  const tasks: Task[] = [];

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEET_CONFIG.tasks.sheetName);

  const { first, schema } = SHEET_CONFIG.tasks;

  for (let row = first.row; ; ++row) {
    if (sheet.getRange(row, first.column).getValue() === '') {
      break;
    }

    const task: Task = {
      contest: {
        id: '',
        title: '',
        problems: [],
      },
      solveTarget: {
        total: 0,
        problems: [],
      },
    };

    for (let column = first.column; column < first.column + schema.length; ++column) {
      const value = sheet.getRange(row, column).getValue();
      switch (schema[column - first.column]) {
        case 'contest.id':
          if (typeof value === 'string') {
            task.contest.id = value;
          } else {
            task.contest.id = value.toString();
          }
          break;
        case 'solveTarget.total':
          if (typeof value === 'number') {
            task.solveTarget.total = value;
          } else {
            task.solveTarget.total = parseInt(value.toString());
          }
          break;
        case 'solveTarget.problems':
          task.solveTarget.problems = (value as string).split(',')
            .map(x => x.trim())
            .filter(x => x.length > 0)
            .map(problemLetterToIndex)
            .sort();
          break;
        default:
          throw new Error(`Unknown schema: ${schema[column - first.column]}`);
      }
    }

    task.contest = fetchContest(task.contest.id);

    Logger.log(task);
    tasks.push(task);
  }

  return tasks;
}


// ==================== Fetching from Vjudge ====================

export function fetchContest(contestId: string): Contest {
  const url = `${CONTESTS_URL}/${contestId}`;
  Logger.log(`Fetching ${url}`);
  const response = UrlFetchApp.fetch(url);
  const html = response.getContentText();
  const json = html.match(/name="dataJson">(.+?)<\/textarea>/)![1];
  const data = JSON.parse(json) as Fetched.ContestPage.Data;
  assert(data.id === parseInt(contestId), 'Contest ID mismatch');

  const problems = data.problems
    .sort((a, b) => (a.num < b.num) ? -1 : (a.num > b.num) ? 1 : 0)
    .map(problem => `${problem.oj}-${problem.probNum}`);

  return {
    id: contestId,
    title: data.title,
    problems,
  };
}

export function fetchSolveSetMapForUser(
  usernames: string[],
  contest: Contest,
): Map<string, Set<number>> {

  // init
  const map = new Map<string, Set<number>>(); // username -> set of solved indices
  for (const username of usernames) {
    map.set(username, new Set<number>());
  }

  const url = `${CONTESTS_URL}/rank/single/${contest.id}`;
  Logger.log(`Fetching ${url}`);
  const response = UrlFetchApp.fetch(url);
  const json = response.getContentText();
  const data = JSON.parse(json) as Fetched.ContestRank.Data;

  const { participants, submissions } = data;
  
  const mapUserIdToIndex = new Map<number, number>();
  
  for (const userIdStr in participants) {
    const [username, _name, _avatarUrl] = participants[userIdStr];
    const index = usernames.indexOf(username);
    if (index !== -1) {
      mapUserIdToIndex.set(parseInt(userIdStr), index);
    }
  }

  for (const submissionTuple of submissions) {
    const [userId, problemIndex, _time, _result] = submissionTuple;
    const index = mapUserIdToIndex.get(userId) ?? -1;
    if (index !== -1) {
      map.get(usernames[index])!.add(problemIndex);
    }
  }

  return map;
}


// ==================== Computations ====================

export function computeProgressMap(
  participants: Participant[],
  tasks: Task[],
): ProgressMap {

  if (participants.length === 0) {
    return new Map();
  }

  const progressMap = new Map<string, ProgressAggregate>();

  const usernames = participants.map(x => x.vjudgeUsernames).flat();
  const solveSetMaps = tasks.map(task => fetchSolveSetMapForUser(usernames, task.contest));
  const solveSetMatrix: Set<number>[][] = [];
  
  for (let i = 0; i < participants.length; ++i) {
    const solveSetList: Set<number>[] = [];
    for (const solveSetMap of solveSetMaps) {
      const solveSet = new Set<number>();
      for (const username of participants[i].vjudgeUsernames) {
        for (const problemIndex of solveSetMap.get(username) ?? []) {
          solveSet.add(problemIndex);
        }
      }
      solveSetList.push(solveSet);
    }
    solveSetMatrix.push(solveSetList);
  }
  
  for (let i = 0; i < solveSetMatrix.length; ++i) {

    const progressAggregate: ProgressAggregate = {
      solveProgressList: [],
      countTasksCompleted: 0,
      countProblemsSolved: 0,
    };

    // (1) solveProgressList

    for (let j = 0; j < solveSetMatrix[i].length; ++j) {
      const task = tasks[j];
      const solveSet = solveSetMatrix[i][j];

      const solveProgress: SolveProgress = {
        count: 0,
        hasProblem: [],
      }

      for (const problemIndex of task.solveTarget.problems) {
        const solved = solveSet.has(problemIndex);
        solveProgress.hasProblem.push(solved);
      }

      const unsolved = solveProgress.hasProblem.filter(x => !x).length;
      solveProgress.count = Math.min(solveSet.size, task.solveTarget.total - unsolved);
      solveProgress.count = Math.max(solveProgress.count, 0);

      progressAggregate.solveProgressList.push(solveProgress);
    }

    // (2) countTasksCompleted

    progressAggregate.countTasksCompleted = progressAggregate.solveProgressList
      .filter((solveProgress, j) =>
        solveProgress.count === tasks[j].solveTarget.total)
      .length;

    // (3) countProblemsSolved

    progressAggregate.countProblemsSolved = progressAggregate.solveProgressList
      .reduce((acc, solveProgress) => acc + solveProgress.count, 0);

    progressMap.set(participants[i].id, progressAggregate);
  }

  return progressMap;
}

export function aggregateTasks(
  tasks: Task[],
): TaskAggregate {

  const taskAggregate: TaskAggregate = {
    tasks,
    totalSolves: 0,
  };

  for (const task of tasks) {
    taskAggregate.totalSolves += task.solveTarget.total;
  }

  return taskAggregate;
}

export function computeRanking(
  participants: Participant[],
  progressMap: ProgressMap,
): RankedParticipant[] {

  const rankedParticipants: RankedParticipant[] = [];

  for (const participant of participants) {
    const progress = progressMap.get(participant.id);
    assert(progress !== undefined, 'Progress not found for participant');
    rankedParticipants.push({
      rank: 0,
      participant,
      progress,
    });
  }

  rankedParticipants.sort((a, b) => {
    const ap = a.progress;
    const bp = b.progress;

    if (ap.countTasksCompleted !== bp.countTasksCompleted) {
      return bp.countTasksCompleted - ap.countTasksCompleted;
    }
    if (ap.countProblemsSolved !== bp.countProblemsSolved) {
      return bp.countProblemsSolved - ap.countProblemsSolved;
    }

    return 0;
  });

  for (let i = 0; i < rankedParticipants.length; ++i) {
    rankedParticipants[i].rank = i + 1;
  }

  return rankedParticipants;
}

// ==================== Writing progress progress to the sheet ====================

export function writeProgressSheet(
  rankedParticipants: RankedParticipant[],
  taskAggreate: TaskAggregate,
) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets().find(sheet =>
    sheet.getName() === SHEET_CONFIG.progress.sheetName
  ) || SpreadsheetApp.getActiveSpreadsheet()
    .insertSheet().setName(SHEET_CONFIG.progress.sheetName);

  // Helpers
  // ----------------
  const makePercentFormatRule = (
    ranges: GoogleAppsScript.Spreadsheet.Range[],
  ) => {
    return SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue(
        COLOR_VALUES.green,
        SpreadsheetApp.InterpolationType.PERCENT,
        '100'
      )
      .setGradientMidpointWithValue(
        COLOR_VALUES.yellow,
        SpreadsheetApp.InterpolationType.PERCENT,
        '75'
      )
      .setGradientMinpointWithValue(
        COLOR_VALUES.red,
        SpreadsheetApp.InterpolationType.PERCENT,
        '50'
      )
      .setRanges(ranges)
      .build();
  }
  const makeStepFormatRules = (
    ranges: GoogleAppsScript.Spreadsheet.Range[],
    maxCount: number,
  ) => {
    return ([
      [1.00, COLOR_VALUES.light_green_2],
      [0.66, COLOR_VALUES.light_yellow_2],
      [0.33, COLOR_VALUES.light_orange_2],
      [0.00, COLOR_VALUES.light_red_2],
    ] as [number, string][])
      .map(([ratio, color]) => {
        return SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThanOrEqualTo(Math.ceil(ratio * maxCount))
          .setBackground(color)
          .setRanges(ranges)
          .build();
      });
  }
  // ----------------

  const columnSpecs: ColumnSpec[] = [
    {
      header: '#',
      width: COLUMN_WIDTHS.xs,
      applyToValueRange: range => {
        range.setBackground(COLOR_VALUES.light_cornflower_blue_3);
      },
      getValueAtRow: rowIndex =>
        rowIndex + 1,
    },
    {
      header: 'Name',
      width: COLUMN_WIDTHS.xl,
      applyToValueRange: range => {
        range.setBackground(COLOR_VALUES.light_cornflower_blue_3);
      },
      getValueAtRow: rowIndex =>
        rankedParticipants[rowIndex].participant.name,
    },
    {
      header: 'Completion',
      applyToHeaderRange: range => {
        range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      },
      subColumns: [
        {
          header: 'Tasks',
          subColumns: [
            {
              header: taskAggreate.tasks.length,
              width: COLUMN_WIDTHS.sm,
              formatRuleHint: taskAggreate.tasks.length,
              applyToValueRange: range => {
                range.setHorizontalAlignment('center')
              },
              getValueAtRow: rowIndex =>
                rankedParticipants[rowIndex].progress.countTasksCompleted,
            },
            {
              header: '100%',
              width: COLUMN_WIDTHS.sm,
              formatRuleHint: '%',
              applyToValueRange: range => {
                range.setHorizontalAlignment('center')
              },
              getValueAtRow: rowIndex =>
                `${Math.round(
                  rankedParticipants[rowIndex].progress.countTasksCompleted /
                  taskAggreate.tasks.length *
                  100)}%`,
            },
          ],
        },
        {
          header: 'Solves',
          subColumns: [
            {
              header: taskAggreate.totalSolves,
              width: COLUMN_WIDTHS.sm,
              formatRuleHint: taskAggreate.totalSolves,
              getValueAtRow: rowIndex =>
                rankedParticipants[rowIndex].progress.countProblemsSolved,
              applyToValueRange: range =>
                range.setHorizontalAlignment('center'),
            },
            {
              header: '100%',
              width: COLUMN_WIDTHS.sm,
              formatRuleHint: '%',
              getValueAtRow: rowIndex =>
                `${Math.round(
                  rankedParticipants[rowIndex].progress.countProblemsSolved /
                  taskAggreate.totalSolves *
                  100)}%`,
              applyToValueRange: range =>
                range.setHorizontalAlignment('center'),
              applyToHeaderRange: range => {
                // freeze up to this row and column
                sheet.setFrozenRows(range.getRow());
                sheet.setFrozenColumns(range.getColumn());
              },
            },
          ],
        },
      ],
    },
  ];

  for (let i = 0; i < taskAggreate.tasks.length; ++i) {
    const task = taskAggreate.tasks[i];

    const titleRichtText = SpreadsheetApp.newRichTextValue()
      .setText(task.contest.title)
      .setLinkUrl(`${CONTESTS_URL}/${task.contest.id}`)
      .build();

    columnSpecs.push({
      header: titleRichtText,
      applyToHeaderRange: range => {
        range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      },
      subColumns: [
        {
          header: 'Count',
          applyToHeaderRange: range => {
            range
              .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
          },
          subColumns: [
            {
              header: task.solveTarget.total,
              width: COLUMN_WIDTHS.sm,
              applyToValueRange: range =>
                range.setHorizontalAlignment('center'),
              getValueAtRow: rowIndex =>
                rankedParticipants[rowIndex].progress.solveProgressList[i].count,
              formatRuleHint: task.solveTarget.total,
            },
          ],
        },
        {
          header: 'Problems',
          applyToHeaderRange: range => {
            range
              .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
          },
          subColumns: task.solveTarget.problems.map((problemIndex, j) => ({
            header: problemIndexToLetter(problemIndex),
            width: COLUMN_WIDTHS.xs,
            getValueAtRow: rowIndex =>
              rankedParticipants[rowIndex].progress
                .solveProgressList[i].hasProblem[j] ? '✅' : '❌',
            applyToValueRange: range =>
              range.setHorizontalAlignment('center'),
          })),
        },
      ]
    });
  }


  function dfsHeaderRows(columnSpec: ColumnSpec): number {
    if ('getValueAtRow' in columnSpec) {
      return 1;
    } else {
      return 1 + columnSpec.subColumns.reduce((acc, subColumn) => {
        return Math.max(acc, dfsHeaderRows(subColumn));
      }, 0);
    }
  }
  const headerRows = columnSpecs.reduce((acc, spec) => {
    return Math.max(acc, dfsHeaderRows(spec));
  }, 0);

  Logger.log(`headerRows: ${headerRows}`);

  // add to format rules array in the process
  const formatRules: GoogleAppsScript.Spreadsheet.ConditionalFormatRule[] = [];

  function dfsWriteColumns(
    row: number,
    column: number,
    columnSpec: ColumnSpec,
  ): number {

    Logger.log(`row: ${row}, column: ${column}, columnSpec: ${JSON.stringify(columnSpec)}`);

    const headerWidth = {
      row: 0,
      column: 0,
    };

    if ('getValueAtRow' in columnSpec) {
      // base case (leaf)

      const valueRange = sheet.getRange(headerRows + 1, column,
        rankedParticipants.length, 1);

      valueRange.setValues(rankedParticipants.map((_, rowIndex) =>
        [columnSpec.getValueAtRow(rowIndex)]));

      sheet.setColumnWidth(column, columnSpec.width);

      if ('formatRuleHint' in columnSpec) {
        const { formatRuleHint } = columnSpec;
        if (formatRuleHint === '%') {
          formatRules.push(makePercentFormatRule([valueRange]));
        } else {
          formatRules.push(...makeStepFormatRules([valueRange], formatRuleHint));
        }
      }

      if ('applyToValueRange' in columnSpec) {
        columnSpec.applyToValueRange(valueRange);
      }

      headerWidth.row = headerRows - row + 1;
      headerWidth.column = 1;

    } else {
      // recursive case (non-leaf)

      const { subColumns } = columnSpec;

      headerWidth.row = 1;
      headerWidth.column = subColumns.reduce((acc, subColumn) => {
        const width = dfsWriteColumns(row + 1, column + acc, subColumn);
        return acc + width;
      }, 0);
    }

    // ----------------

    if (headerWidth.column > 0) {
      const headerRange = sheet.getRange(row, column, headerWidth.row, headerWidth.column);
      headerRange.setBackground(row == 2
        ? COLOR_VALUES.light_cornflower_blue_3
        : COLOR_VALUES.light_grey_2
      );
      headerRange.setHorizontalAlignment('center');
      headerRange.setVerticalAlignment('top');

      if (typeof columnSpec.header == 'string' ||
        typeof columnSpec.header == 'number') {
        headerRange.setValue(columnSpec.header);
      } else {
        headerRange.setRichTextValue(columnSpec.header as
          GoogleAppsScript.Spreadsheet.RichTextValue);
      }

      if ('applyToHeaderRange' in columnSpec) {
        columnSpec.applyToHeaderRange(headerRange);
      }

      headerRange.merge();
    }

    return headerWidth.column;
  }

  sheet.clearFormats();
  sheet.getDataRange().breakApart();

  const totalColumns = columnSpecs.reduce((acc, spec) => {
    return acc + dfsWriteColumns(1, acc, spec);
  }, 1);

  sheet.setRowHeightsForced(1, 1, 200);

  const totalRows = rankedParticipants.length + headerRows;

  sheet.setConditionalFormatRules(formatRules);

  // clear the rest of the sheet
  const dataRangeRows = sheet.getLastRow();
  const dataRangeColumns = sheet.getLastColumn();
  if (dataRangeRows > totalRows) {
    sheet.getRange(totalRows + 1, 1, dataRangeRows - headerRows, dataRangeColumns)
      .clear();
  }
  if (dataRangeColumns > totalColumns) {
    sheet.getRange(1, totalColumns + 1, dataRangeRows, dataRangeColumns - totalColumns)
      .clear();
  }
}


// ==================== Helpers ====================

function assert(condition: boolean, message: string) {
  if (!condition) {
    throw new Error(message);
  }
}

export function problemIndexToLetter(index: number) {
  if (!Number.isInteger(index)) {
    throw new Error('Index must be an integer');
  }
  if (index < 0) {
    throw new Error('Index must be non-negative');
  }
  if (index >= 702) {
    throw new Error('Index must be less than 702');
  }
  const asciiOfA = 'A'.charCodeAt(0);
  if (index < 26) {
    return String.fromCharCode(asciiOfA + index);
  } else {
    return String.fromCharCode(asciiOfA + Math.floor(index / 26) - 1)
      + String.fromCharCode(asciiOfA + index % 26);
  }
}

export function problemLetterToIndex(letter: string) {
  if (typeof letter !== 'string') {
    throw new Error('Letter must be a string');
  }
  if (letter.length === 0) {
    throw new Error('Letter must not be empty');
  }
  if (!/^[a-zA-Z]+$/.test(letter)) {
    throw new Error('Letter must contain only alphabetic characters');
  }
  const asciiOfA = 'A'.charCodeAt(0);
  letter = letter.toUpperCase();
  if (letter.length === 1) {
    return letter.charCodeAt(0) - asciiOfA;
  } else {
    return (letter.charCodeAt(0) - asciiOfA + 1) * 26
      + (letter.charCodeAt(1) - asciiOfA);
  }
}
