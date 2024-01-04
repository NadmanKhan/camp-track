// Webpack GasPlugin will turn this into a dummy top level function for GAS:
global.main = () => {};

// ==================== Constants ====================

const CONTESTS_URL = `https://vjudge.net/contest`;

const PREVIOUS_RUN_DATESTRING =
    PropertiesService.getScriptProperties().getProperty('lastRun');
const NOW = new Date();

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
    summary: {
        sheetName: 'Summary',
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
    xl: 200,
} as const;

// ==================== Main function ====================

(function main() {
    PropertiesService.getScriptProperties().setProperty(
        'lastRun',
        NOW.toString()
    );

    const participants = readParticipants();

    const tasks = readTasks();

    const progressMap = computeProgressMap(participants, tasks);

    const rankedParticipants = computeRanking(participants, progressMap);

    const taskAggreate = aggregateTasks(tasks);

    writeProgressSheet(rankedParticipants, taskAggreate);

    writeSummarySheet(participants, progressMap, taskAggreate);
})();

// ==================== Read from sheet ====================

export function readParticipants(): Participant[] {
    const participants: Participant[] = [];

    const { sheetName, first, schema } = SHEET_CONFIG.participants;

    const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

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

        for (
            let column = first.column;
            column < first.column + schema.length;
            ++column
        ) {
            const value = sheet.getRange(row, column).getValue();
            switch (schema[column - first.column]) {
                case 'id':
                case 'name':
                case 'email':
                    if (typeof value === 'string') {
                        participant[schema[column - first.column]] = value;
                    } else {
                        participant[schema[column - first.column]] =
                            value.toString();
                    }
                    break;
                case 'vjudgeUsernames':
                    participant.vjudgeUsernames = (value as string)
                        .split(',')
                        .map(x => x.trim())
                        .filter(x => x.length > 0)
                        .sort();
                    break;
                default:
                    throw new Error(
                        `Unknown schema: ${schema[column - first.column]}`
                    );
            }
        }

        Logger.log(participant);
        participants.push(participant);
    }

    return participants;
}

export function readTasks(): Task[] {
    const tasks: Task[] = [];

    const { sheetName, first, schema } = SHEET_CONFIG.tasks;

    const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

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

        for (
            let column = first.column;
            column < first.column + schema.length;
            ++column
        ) {
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
                    task.solveTarget.problems = (value as string)
                        .split(',')
                        .map(x => x.trim())
                        .filter(x => x.length > 0)
                        .map(problemLetterToIndex)
                        .sort();
                    break;
                default:
                    throw new Error(
                        `Unknown schema: ${schema[column - first.column]}`
                    );
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
        .sort((a, b) => (a.num < b.num ? -1 : a.num > b.num ? 1 : 0))
        .map(problem => `${problem.oj}-${problem.probNum}`);

    return {
        id: contestId,
        title: data.title,
        problems,
    };
}

export function fetchSolveSetMapForUser(
    usernames: string[],
    contest: Contest
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
        const [userId, problemIndex, result, _time] = submissionTuple;
        const index = mapUserIdToIndex.get(userId) ?? -1;
        if (index !== -1 && result === 1) {
            map.get(usernames[index])!.add(problemIndex);
        }
    }

    const asArray = Array.from(map.entries()).map(([username, set]) => [
        username,
        Array.from(set).toSorted((a, b) => a - b),
    ]);
    Logger.log(`Solve set map for ${contest.id}: ${JSON.stringify(asArray)}`);

    return map;
}

// ==================== Computations ====================

export function computeProgressMap(
    participants: Participant[],
    tasks: Task[]
): ProgressMap {
    if (participants.length === 0) {
        return new Map();
    }

    const progressMap = new Map<string, ProgressAggregate>();

    const usernames = participants.map(x => x.vjudgeUsernames).flat();
    const solveSetMaps = tasks.map(task =>
        fetchSolveSetMapForUser(usernames, task.contest)
    );
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
            };

            for (const problemIndex of task.solveTarget.problems) {
                const solved = solveSet.has(problemIndex);
                solveProgress.hasProblem.push(solved);
            }

            const unsolved = solveProgress.hasProblem.filter(x => !x).length;
            solveProgress.count = Math.min(
                solveSet.size,
                task.solveTarget.total - unsolved
            );
            solveProgress.count = Math.max(solveProgress.count, 0);

            Logger.log(`For ${participants[i].name} on ${task.contest.id} ---`);
            Logger.log(`solve set: ${Array.from(solveSet)}`)
            Logger.log(`progress : ${JSON.stringify(solveProgress)}`);
            progressAggregate.solveProgressList.push(solveProgress);
        }

        // (2) countTasksCompleted

        progressAggregate.countTasksCompleted =
            progressAggregate.solveProgressList.filter(
                (solveProgress, j) =>
                    solveProgress.count === tasks[j].solveTarget.total
            ).length;

        // (3) countProblemsSolved

        progressAggregate.countProblemsSolved =
            progressAggregate.solveProgressList.reduce(
                (acc, solveProgress) => acc + solveProgress.count,
                0
            );
        
        Logger.log(`progress aggregate for ${participants[i].name}: ${JSON.stringify(progressAggregate)}`);
        progressMap.set(participants[i].id, progressAggregate);
    }

    return progressMap;
}

export function aggregateTasks(tasks: Task[]): TaskAggregate {
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
    progressMap: ProgressMap
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

export function writeColumns(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    columnWriteConfigs: ColumnWriteConfig[],
    numValueRows: number,
    columnOffset = 0
): {
    numHeaderRows: number;
    numColumns: number;
    newFormatRules: GoogleAppsScript.Spreadsheet.ConditionalFormatRule[];
} {
    const { numHeaderRows, numColumns } = getNumWritable(columnWriteConfigs);
    const formatRules: GoogleAppsScript.Spreadsheet.ConditionalFormatRule[] =
        [];

    function dfsWriteColumns(
        row: number,
        column: number,
        columnConfig: ColumnWriteConfig
    ): number {
        Logger.log(
            `row: ${row}, column: ${column}, columnConfig: ${JSON.stringify(
                columnConfig
            )}`
        );

        const headerWidth = {
            row: 0,
            column: 0,
        };

        if ('getValueAtRow' in columnConfig) {
            // base case (leaf)

            const valueRange = sheet.getRange(
                numHeaderRows + 1,
                column,
                numValueRows,
                1
            );

            // (1) values
            for (let i = 0; i < numValueRows; ++i) {
                const value = columnConfig.getValueAtRow(i);
                setCellValue(sheet, numHeaderRows + i + 1, column, value);
            }

            // (2) column width
            sheet.setColumnWidth(column, columnConfig.width);

            // (3) format rules
            if ('formatRuleHint' in columnConfig) {
                const { formatRuleHint } = columnConfig;
                if (formatRuleHint === '%') {
                    formatRules.push(makePercentFormatRule([valueRange]));
                } else {
                    formatRules.push(
                        ...makeStepFormatRules([valueRange], formatRuleHint)
                    );
                }
            }

            // (4) applyToValueRange
            if ('applyToValueRange' in columnConfig) {
                columnConfig.applyToValueRange(valueRange);
            }

            headerWidth.row = numHeaderRows - row + 1;
            headerWidth.column = 1;
        } else {
            // recursive case (non-leaf)

            const { subColumns } = columnConfig;

            headerWidth.row = 1;
            headerWidth.column = subColumns.reduce((acc, subColumn) => {
                const width = dfsWriteColumns(row + 1, column + acc, subColumn);
                return acc + width;
            }, 0);
        }

        // ----------------

        if (headerWidth.column > 0) {
            const headerRange = sheet.getRange(
                row,
                column,
                headerWidth.row,
                headerWidth.column
            );

            setCellValueAtRange(headerRange, columnConfig.header);

            headerRange
                .setBackground(
                    row % 2 == 0
                        ? COLOR_VALUES.light_cornflower_blue_3
                        : COLOR_VALUES.light_grey_2
                )
                .setHorizontalAlignment('center')
                .setVerticalAlignment('top')
                .merge();

            if ('applyToHeaderRange' in columnConfig) {
                columnConfig.applyToHeaderRange(headerRange);
            }
        }

        return headerWidth.column;
    }

    const numColumns2 =
        -1 +
        columnWriteConfigs.reduce((acc, conf) => {
            return acc + dfsWriteColumns(1, columnOffset + acc, conf);
        }, 1);

    assert(
        numColumns === numColumns2,
        `numColumns mismatch: ${numColumns} !== ${numColumns2}; fix algorithm`
    );

    return {
        numHeaderRows,
        numColumns,
        newFormatRules: formatRules,
    };
}

export function writeProgressSheet(
    rankedParticipants: RankedParticipant[],
    taskAggreate: TaskAggregate
) {
    const { sheetName } = SHEET_CONFIG.progress;

    const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

    const columnWriteConfigs: ColumnWriteConfig[] = [
        {
            header: '#',
            width: COLUMN_WIDTHS.xs,
            applyToValueRange: range => {
                range.setBackground(COLOR_VALUES.light_cornflower_blue_3);
            },
            getValueAtRow: rowIndex => rowIndex + 1,
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
                                range.setHorizontalAlignment('center');
                            },
                            getValueAtRow: rowIndex =>
                                rankedParticipants[rowIndex].progress
                                    .countTasksCompleted,
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
                                rankedParticipants[rowIndex].progress
                                    .countProblemsSolved,
                            applyToValueRange: range =>
                                range.setHorizontalAlignment('center'),
                        },
                        {
                            header: '100%',
                            width: COLUMN_WIDTHS.sm,
                            formatRuleHint: '%',
                            getValueAtRow: rowIndex =>
                                `${Math.round(
                                    (rankedParticipants[rowIndex].progress
                                        .countProblemsSolved /
                                        taskAggreate.totalSolves) *
                                        100
                                )}%`,
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

        columnWriteConfigs.push({
            header: titleRichtText,
            applyToHeaderRange: range => {
                range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
            },
            subColumns: [
                {
                    header: 'Count',
                    applyToHeaderRange: range => {
                        range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
                    },
                    subColumns: [
                        {
                            header: task.solveTarget.total,
                            width: COLUMN_WIDTHS.sm,
                            applyToValueRange: range =>
                                range.setHorizontalAlignment('center'),
                            getValueAtRow: rowIndex =>
                                rankedParticipants[rowIndex].progress
                                    .solveProgressList[i].count,
                            formatRuleHint: task.solveTarget.total,
                        },
                    ],
                },
                {
                    header: 'Problems',
                    applyToHeaderRange: range => {
                        range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
                    },
                    subColumns: task.solveTarget.problems.map(
                        (problemIndex, j) => ({
                            header: problemIndexToLetter(problemIndex),
                            width: COLUMN_WIDTHS.xs,
                            getValueAtRow: rowIndex =>
                                rankedParticipants[rowIndex].progress
                                    .solveProgressList[i].hasProblem[j]
                                    ? '✅'
                                    : '❌',
                            applyToValueRange: range =>
                                range.setHorizontalAlignment('center'),
                        })
                    ),
                },
            ],
        });
    }

    sheet.clearFormats();
    sheet.getDataRange().breakApart();

    const { numHeaderRows, numColumns, newFormatRules } = writeColumns(
        sheet,
        columnWriteConfigs,
        rankedParticipants.length
    );
    const numRows = rankedParticipants.length + numHeaderRows;

    sheet.clearConditionalFormatRules();
    sheet.setConditionalFormatRules(newFormatRules);
    sheet.setRowHeightsForced(1, 1, 200);

    // clear the rest of the sheet
    const dataRangeRows = sheet.getLastRow();
    const dataRangeColumns = sheet.getLastColumn();
    if (dataRangeRows > numRows) {
        sheet
            .getRange(
                numRows + 1,
                1,
                dataRangeRows - numHeaderRows,
                dataRangeColumns
            )
            .clear();
    }
    if (dataRangeColumns > numColumns) {
        sheet
            .getRange(
                1,
                numColumns + 1,
                dataRangeRows,
                dataRangeColumns - numColumns
            )
            .clear();
    }
}

export function writeSummarySheet(
    participants: Participant[],
    progressMap: ProgressMap,
    taskAggreate: TaskAggregate
) {
    if (PREVIOUS_RUN_DATESTRING === null) {
        return;
    }
    const previousRun = new Date(PREVIOUS_RUN_DATESTRING);
    if (
        previousRun.getMonth() === NOW.getMonth() &&
        previousRun.getDate() === NOW.getDate()
    ) {
        return;
    }

    const { sheetName } = SHEET_CONFIG.summary;

    const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

    const orderedParticipants = [...participants].sort((a, b) =>
        a.name.localeCompare(b.name)
    );

    if (sheet.getRange(1, 1).getValue() !== 'Name') {
        sheet.clearContents().clearFormats();
        [
            { header: 'Name', width: COLUMN_WIDTHS.xl },
            { header: 'ID', width: COLUMN_WIDTHS.md },
        ].forEach((conf, i) => {
            sheet
                .getRange(1, i + 1, 3 /* numHeaderRows */, 1)
                .setValue(conf.header)
                .merge()
                .setBackground(COLOR_VALUES.light_grey_3)
                .setHorizontalAlignment('center')
                .setVerticalAlignment('top');
            sheet.setColumnWidth(i + 1, conf.width);
        });
    }

    const columnWriteConfigs: ColumnWriteConfig[] = [
        {
            header: NOW.toLocaleString('en-GB', { timeZone: 'Asia/Dhaka' }),
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
                            getValueAtRow: rowIndex =>
                                progressMap.get(
                                    orderedParticipants[rowIndex].id
                                )!.countTasksCompleted,
                            applyToValueRange: range =>
                                range.setHorizontalAlignment('center'),
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
                                progressMap.get(
                                    orderedParticipants[rowIndex].id
                                )!.countProblemsSolved,
                            applyToValueRange: range =>
                                range.setHorizontalAlignment('center'),
                        },
                        {
                            header: '100%',
                            width: COLUMN_WIDTHS.sm,
                            formatRuleHint: '%',
                            getValueAtRow: rowIndex =>
                                `${Math.round(
                                    (progressMap.get(
                                        orderedParticipants[rowIndex].id
                                    )!.countProblemsSolved /
                                        taskAggreate.totalSolves) *
                                        100
                                )}%`,
                            applyToValueRange: range =>
                                range.setHorizontalAlignment('center'),
                        },
                    ],
                },
            ],
        },
    ];

    const { numHeaderRows, numColumns } = getNumWritable(columnWriteConfigs);
    assert(numHeaderRows === 3, 'numHeaderRows must be 3; fix algorithm');
    assert(numColumns == 3, 'numColumns must be 3; fix algorithm');

    sheet.setFrozenColumns(2);
    sheet.setFrozenRows(numHeaderRows);

    const isWrittenIdSet = new Set<string>();

    // filter from existing rows
    let row = numHeaderRows + 1;
    for (; row <= sheet.getLastRow(); ++row) {
        const name = sheet.getRange(row, 1).getValue();
        const idValue = sheet.getRange(row, 2).getValue();
        const id = typeof idValue === 'number' ? idValue.toString() : idValue;
        if (name === '' || id === '' || !progressMap.has(id)) {
            // participant not found, delete this row
            Logger.log(`id ${id} not found; deleting row ${row}`);
            sheet.deleteRow(row);
            --row;
            continue;
        }
        isWrittenIdSet.add(id);
    }
    // append remaining rows
    for (const participant of orderedParticipants) {
        if (isWrittenIdSet.has(participant.id)) {
            continue;
        }
        Logger.log(`id ${participant.id} not added; appending row`);
        // do not use sheet.appendRow because we have merged rows above
        sheet.getRange(row, 1, 1, 2).setValues([
            [participant.name, participant.id],
        ]);
        ++row;
    }
    // sort rows
    sheet
        .getRange(numHeaderRows + 1, 1, sheet.getLastRow() - numHeaderRows, 2)
        .sort(1);

    // write columns
    sheet.insertColumnsAfter(2, numColumns);
    const {
        numHeaderRows: numHeaderRows2,
        numColumns: numColumns2,
        newFormatRules,
    } = writeColumns(
        sheet,
        columnWriteConfigs,
        orderedParticipants.length,
        2 /* columnOffset */
    );

    assert(
        numHeaderRows === numHeaderRows2,
        `numHeaderRows mismatch: ${numHeaderRows} !== ${numHeaderRows2}; fix algorithm`
    );
    assert(
        numColumns === numColumns2,
        `numColumns mismatch: ${numColumns} !== ${numColumns2}; fix algorithm`
    );
    sheet.setConditionalFormatRules([
        ...sheet.getConditionalFormatRules(),
        ...newFormatRules,
    ]);
}

// ==================== Helpers ====================

export function assert(condition: boolean, message?: string) {
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
        return (
            String.fromCharCode(asciiOfA + Math.floor(index / 26) - 1) +
            String.fromCharCode(asciiOfA + (index % 26))
        );
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
        return (
            (letter.charCodeAt(0) - asciiOfA + 1) * 26 +
            (letter.charCodeAt(1) - asciiOfA)
        );
    }
}

export function makePercentFormatRule(
    ranges: GoogleAppsScript.Spreadsheet.Range[]
) {
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

export function makeStepFormatRules(
    ranges: GoogleAppsScript.Spreadsheet.Range[],
    maxCount: number
) {
    return (
        [
            [1.0, COLOR_VALUES.light_green_2],
            [0.66, COLOR_VALUES.light_yellow_2],
            [0.33, COLOR_VALUES.light_orange_2],
            [0.0, COLOR_VALUES.light_red_2],
        ] as [number, string][]
    ).map(([ratio, color]) => {
        return SpreadsheetApp.newConditionalFormatRule()
            .whenNumberGreaterThanOrEqualTo(Math.ceil(ratio * maxCount))
            .setBackground(color)
            .setRanges(ranges)
            .build();
    });
}

export function setCellValue(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    row: number,
    column: number,
    value: CellValue
) {
    const range = sheet.getRange(row, column);
    setCellValueAtRange(range, value);
}

export function setCellValueAtRange(
    range: GoogleAppsScript.Spreadsheet.Range,
    value: CellValue
) {
    range.clearContent().clearFormat();
    if (typeof value === 'string' || typeof value === 'number') {
        range.setValue(value);
    } else {
        range.setRichTextValue(value);
    }
}

export function getNumWritable(columnWriteConfigs: ColumnWriteConfig[]): {
    numHeaderRows: number;
    numColumns: number;
} {
    function dfsNumHeaderRows(columnConfig: ColumnWriteConfig): number {
        if ('getValueAtRow' in columnConfig) {
            return 1;
        } else {
            return (
                1 +
                columnConfig.subColumns.reduce(
                    (acc, subColumn) =>
                        Math.max(acc, dfsNumHeaderRows(subColumn)),
                    0
                )
            );
        }
    }
    function dfsNumColumns(columnConfig: ColumnWriteConfig): number {
        if ('getValueAtRow' in columnConfig) {
            return 1;
        } else {
            return columnConfig.subColumns.reduce(
                (acc, subColumn) => acc + dfsNumColumns(subColumn),
                0
            );
        }
    }
    return {
        numHeaderRows: columnWriteConfigs.reduce(
            (acc, conf) => Math.max(acc, dfsNumHeaderRows(conf)),
            0
        ),
        numColumns: columnWriteConfigs.reduce(
            (acc, conf) => acc + dfsNumColumns(conf),
            0
        ),
    };
}
