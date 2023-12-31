declare namespace Fetched {
    namespace ContestPage {
        /**
         * @link https://vjudge.net/contest/{CONTEST_ID}
         * (e.g. https://vjudge.net/contest/580361)
         * (at \<textarea name="dataJson">)
         */
        type Data = {
            id: number;
            title: string;
            type: number;
            openness: number;
            authStatus: number;
            begin: number;
            end: number;
            createTime: number;
            started: boolean;
            ended: boolean;
            managerId: number;
            managerName: string;
            fav: boolean;
            description: Description;
            announcement: string;
            problems: Problem[];
            problemsHash: string;
            privatePeerContestIds: any[];
            enableTimeMachine: boolean;
            sumTime: boolean;
            penalty: number;
            partialScore: boolean;
            customizedWeight: boolean;
            showPeers: boolean;
            clonedFromId: number;
            clonedFromTitle: string;
        };

        type Description = {
            format: string;
            content: string;
        };

        type Problem = {
            pid: number;
            title: string;
            oj: string;
            probNum: string;
            num: string;
            publicDescId: number;
            publicDescVersion: number;
            properties: Property[];
            weight: number;
            languages: { [key: string]: string };
            submitMethods: boolean[];
            submitTips?: string;
        };

        type Property = {
            title: string;
            content: string;
            spoiler: boolean;
        };
    }

    /**
     * @link https://vjudge.net/contest/rank/single/{CONTEST_ID}
     * (e.g. https://vjudge.net/contest/rank/single/580361)
     */
    namespace ContestRank {
        type Data = {
            id: number;
            title: string;
            begin: number; // unix timestamp in ms
            length: number; // in ms
            isReplay: boolean;
            participants: { [key: string]: ParticipantTuple }; // key is the user ID
            submissions: SubmissionTuple[];
        };
        /**
         * [<handle>, <name>, <avatar URL>]
         */
        type ParticipantTuple = [string, string, string];
        /**
         * A submission is given in an array of 4 elements:
         * 0. User ID
         * 1. Problem index (A = 0, B = 1, ...)
         * 2. Result (0 for Wrong Answer, 1 for Accepted)
         * 3. Time since contest beginning (in seconds)
         */
        type SubmissionTuple = [number, number, number, number];
    }

    /**
     * @link https://vjudge.net/status/data?draw={DRAW}&start={START}&length={LENGTH}&un={USERNAME}&OJId={OJ}&probNum={PROBLEM_ID_IN_OJ}&res={RESULT:0|1}&language={LANGUAGE}&onlyFollowee={:true|false}&_={SYSTEM_TIME_MS} (e.g. https://vjudge.net/status/data?draw=1&start=0&length=20&un=NadmanKhan&OJId=All&probNum=&res=1&language=&onlyFollowee=false&_=1703520299520)
     * * The `start` parameter is the index of the first submission, starting from 0.
     * * The `length` parameter is the number of submissions to fetch, no more than 20.
     * * Only the last 200 submissions in total are available by pagination.
     */
    namespace UserStatus {
        type Data = {
            data: Submission[];
            recordsTotal: number;
            recordsFiltered: number;
            draw: number;
        };

        type Submission = {
            memory: number;
            access: number;
            statusType: number;
            avatarUrl: string;
            runtime: number;
            contestOpenness: number;
            language: string;
            userName: string;
            userId: number;
            languageCanonical: string;
            contestId: number;
            contestNum: string;
            processing: boolean;
            runId: number;
            time: number;
            oj: string;
            problemId: number;
            sourceLength: number;
            probNum: string;
            status: string;
        };
    }
}

type Contest = {
    id: string;
    title: string;
    problems: string[];
};

type Participant = {
    id: string;
    name: string;
    email: string;
    vjudgeUsernames: string[];
};

type SolveTarget = {
    total: number;
    problems: number[];
};

type SolveProgress = {
    count: number;
    hasProblem: boolean[];
};

type Task = {
    contest: Contest;
    solveTarget: SolveTarget;
};

type TaskAggregate = {
    tasks: Task[];
    totalSolves: number;
};

type ProgressAggregate = {
    solveProgressList: SolveProgress[];
    countTasksCompleted: number;
    countProblemsSolved: number;
};

type ProgressMap = Map<string, ProgressAggregate>;

type RankedParticipant = {
    rank: number;
    participant: Participant;
    progress: ProgressAggregate;
};

type SheetConfig = {
    [key: string]: {
        sheetName: string;
        first?: {
            row: number;
            column: number;
        };
        schema?: string[];
    };
};

type CellValue = string | number | GoogleAppsScript.Spreadsheet.RichTextValue;

type ColumnWriteConfig = {
    header: CellValue;
} & (
    | {
          width: number;
          formatRuleHint?: '%' | number;
          getValueAtRow: (rowIndex: number) => CellValue;
          applyToHeaderRange?: (
              range: GoogleAppsScript.Spreadsheet.Range
          ) => void;
          applyToValueRange?: (
              range: GoogleAppsScript.Spreadsheet.Range
          ) => void;
      }
    | {
          subColumns: ColumnWriteConfig[];
      }
);
