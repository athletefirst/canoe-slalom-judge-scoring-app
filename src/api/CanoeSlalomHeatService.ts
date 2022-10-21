import CanoeSlalomHeatData from "../dao/CanoeSlalomHeatData";

namespace CanoeSlalomHeatService {

    namespace CONSTS {
        export const DATA_HEADER_ROW1 = 1
        export const DATA_HEADER_ROW2 = 2
        /**
         * データ開始行
         */
        export const DATA_TOP_ROW = 3;
        /**
         * 選手データの開始列
         */
        export const RUNNER_COLUMN = 1;
        /**
         * [0] bib
         * [1] heat
         * [2] isLocked
         */
        export const RUNNER_LENGTH = 3;
        /**
         * スタートタイムの開始列
         */
        export const START_COLUMN = 4;
        /**
         * [0] hours
         * [1] minutes
         * [2] seconds
         * [3] judge
         */
        export const START_LENGTH = 4;
        /**
         * フィニッシュタイムの開始列
         */
        export const FINISH_COLUMN = 8;
        /**
         * [0] hours
         * [1] minutes
         * [2] seconds
         * [3] judge
         */
        export const FINISH_LENGTH = 4;
        /**
         * ゲート判定の開始列
         */
        export const GATE_COLUMN = 12;
    }

    export function createNewSheet(heatName: string) {
        const spread = SpreadsheetApp.getActiveSpreadsheet();
        const sheetName = `heat:${heatName}`
        const sheet = spread.insertSheet(sheetName, spread.getSheets().length);
        // 選手データ
        sheet.getRange(CONSTS.DATA_HEADER_ROW1, CONSTS.RUNNER_COLUMN).setValue('RUNNER');
        sheet.getRange(CONSTS.DATA_HEADER_ROW2, CONSTS.RUNNER_COLUMN, 1, CONSTS.RUNNER_LENGTH).setValues([
            [
                'bib',
                'tag',
                'locked',
            ]
        ]);
        // スタートタイム
        sheet.getRange(CONSTS.DATA_HEADER_ROW1, CONSTS.START_COLUMN).setValue('START');
        sheet.getRange(CONSTS.DATA_HEADER_ROW2, CONSTS.START_COLUMN, 1, CONSTS.START_LENGTH).setValues([
            [
                'hours',
                'minutes',
                'seconds',
                'judge',
            ]
        ]);
        // フィニッシュタイム
        sheet.getRange(CONSTS.DATA_HEADER_ROW1, CONSTS.FINISH_COLUMN).setValue('FINISH');
        sheet.getRange(CONSTS.DATA_HEADER_ROW2, CONSTS.FINISH_COLUMN, 1, CONSTS.FINISH_LENGTH).setValues([
            [
                'hours',
                'minutes',
                'seconds',
                'judge',
            ]
        ]);
        // ゲート判定
        sheet.getRange(CONSTS.DATA_HEADER_ROW1, CONSTS.GATE_COLUMN, 1, CanoeSlalomHeatData.CONSTS.GATE_MAX).setValues([
            [...Array(CanoeSlalomHeatData.CONSTS.GATE_MAX)].map((_, i) => `G[${i + 1}]`)
        ]);
        // スクロール固定
        sheet.setFrozenRows(CONSTS.DATA_HEADER_ROW2);
        sheet.setFrozenColumns(CONSTS.RUNNER_LENGTH);
        return sheet;
    }

    export function putGateSettings(heatName: string, gateSettings: CanoeSlalomHeatData.gateSetting[]) {
        const sheet = getSheetByHeatName(heatName);
        // 検証
        gateSettings.forEach(setting => {
            CanoeSlalomHeatData.validateGateNum(setting.num);
        });
        // 設定
        const range = sheet.getRange(CONSTS.DATA_HEADER_ROW2, CONSTS.GATE_COLUMN, 1, CanoeSlalomHeatData.CONSTS.GATE_MAX);
        const gateTypes = CanoeSlalomHeatData.convGateTypeList(range.getValues()[0]);
        gateSettings.forEach(setting => {
            gateTypes[(setting.num - 1)] = CanoeSlalomHeatData.convGateType(setting.direction);
        });
        range.setValues([CanoeSlalomHeatData.toStringGateTypeList(gateTypes)]);
    }

    export function putRunners(heatName: string, runners: CanoeSlalomHeatData.runner[]) {
        const sheet = getSheetByHeatName(heatName);
        // 検証
        runners.forEach(runner => {
            if (runner.row < 0) {
                throw new Error(`Invalid row: ${runner.row}`);
            }
        });
        // 設定
        runners.forEach(runner => {
            const sheetRow = CONSTS.DATA_TOP_ROW + runner.row;
            const range = sheet.getRange(sheetRow, CONSTS.RUNNER_COLUMN, 1, CONSTS.RUNNER_LENGTH);
            const values = [
                runner.bib,
                runner.tag,
                runner.locked ? runner.locked : '',
            ];
            range.setValues([values]);
        });
    }

    /**
     * データセット取得条件
     */
    export type Criteria = {
        /**
         * ヒート名（シート名）
         */
        heatName: string;
        /**
         * スタートタイムの要求
         */
        started?: boolean;
        /**
         * ゴールタイムの要求
         */
        finished?: boolean;
        /**
         * ゲート判定の要求
         */
        gates?: {
            /**
             * 開始ゲート番号
             */
            beginGate: number;
            /**
             * 要求ゲート数
             */
            gateLength: number;
        };
    }

    /**
     * スプレッドシートの取得
     * @param heatName ヒート名
     * @returns スプレッドシート
     */
    function getSheetByHeatName(heatName: string) {
        const sheetName = `heat:${heatName}`;
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        if (!sheet) {
            throw new Error(`Missing sheet: '${sheetName}'`);
        }
        return sheet;
    }

    export function getHeatNameList() {
        const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
        const r: string[] = [];
        sheets.forEach(v => {
            const sheetName = v.getSheetName();
            if (sheetName.startsWith('heat:')) {
                r.push(sheetName.substring(5));
            }
        });
        return r;
    }

    /**
     * データ件数の取得
     * @param sheet スプレッドシート
     * @returns 1以上
     */
    function getRowCount(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
        const rowCount = sheet.getLastRow() - (CONSTS.DATA_TOP_ROW - 1);
        if (rowCount < 1) {
            throw new Error(`No records.`);
        }
        return rowCount;
    }

    /**
     * データセットの取得
     * @param criteria 取得条件
     * @returns データセット
     */
    export function getDataset(criteria: Criteria) {
        const sheet = getSheetByHeatName(criteria.heatName);
        const rowCount = getRowCount(sheet);
        const dataset: CanoeSlalomHeatData.Dataset = {
            heatName: criteria.heatName,
            runs: [],
        };
        const rsLocked: boolean[] = []; // isLockedの一時リスト

        (() => {
            // 選手データリストの取得
            const rs = sheet.getRange(CONSTS.DATA_TOP_ROW, CONSTS.RUNNER_COLUMN, rowCount, CONSTS.RUNNER_LENGTH).getValues();
            rs.forEach((r, i) => {
                const runner: CanoeSlalomHeatData.runner = {
                    row: i,
                    bib: String(r[0]),
                    tag: String(r[1]),
                };
                const run: CanoeSlalomHeatData.run = {
                    runner,
                };
                const isLocked = String(r[2]) != '';
                if (isLocked) {
                    runner.locked = String(r[2]);
                }
                rsLocked.push(isLocked);  // isLocked
                dataset.runs.push(run);
            });
        })();

        if (criteria.started) {
            // スタートタイム
            const rs = sheet.getRange(CONSTS.DATA_TOP_ROW, CONSTS.START_COLUMN, rowCount, CONSTS.START_LENGTH).getValues();
            rs.forEach((r, i) => {
                const run = dataset.runs[i];
                const hms = {
                    hours: r[0],
                    minutes: r[1],
                    seconds: r[2],
                }
                const seconds = CanoeSlalomHeatData.hmsToSeconds(hms);
                const judge = CanoeSlalomHeatData.validateStartedTimeJudge(r[3]);
                run.started = {
                    seconds,
                    judge,
                    fetching: {},
                };
                if (rsLocked[i]) {
                    // isLocked
                    run.started.isLocked = true;
                }
            });
        }

        if (criteria.finished) {
            // フィニッシュタイム
            const rs = sheet.getRange(CONSTS.DATA_TOP_ROW, CONSTS.FINISH_COLUMN, rowCount, CONSTS.FINISH_LENGTH).getValues();
            rs.forEach((r, i) => {
                const run = dataset.runs[i];
                const hms = {
                    hours: r[0],
                    minutes: r[1],
                    seconds: r[2],
                }
                const seconds = CanoeSlalomHeatData.hmsToSeconds(hms);
                const judge = CanoeSlalomHeatData.validateFinishedTimeJudge(r[3]);
                run.finished = {
                    seconds,
                    judge,
                    fetching: {},
                };
                if (rsLocked[i]) {
                    // isLocked
                    run.finished.isLocked = true;
                }
            });
        }

        if (criteria.gates) {
            // ゲート判定
            const beginGate = criteria.gates.beginGate;
            const gateLength = criteria.gates.gateLength;
            try {
                CanoeSlalomHeatData.validateGateNum(beginGate);
                if (gateLength < 1) {
                    throw new Error(`Invalid gateLength: ${gateLength}`);
                }
                CanoeSlalomHeatData.validateGateNum(beginGate + gateLength - 1);
            } catch (error) {
                throw new Error(`Invalid gates: ${beginGate} (${gateLength})`);
            }
            const column = CONSTS.GATE_COLUMN + beginGate - 1;
            const gateTypes = CanoeSlalomHeatData.convGateTypeList(sheet.getRange(CONSTS.DATA_HEADER_ROW2, column, 1, gateLength).getValues()[0]);
            const rs = sheet.getRange(CONSTS.DATA_TOP_ROW, column, rowCount, gateLength).getValues();
            rs.forEach((r, i) => {
                const gates: CanoeSlalomHeatData.gate[] = [];
                r.forEach((g, j) => {
                    const num = beginGate + j;
                    const direction = gateTypes[j];
                    const judge = CanoeSlalomHeatData.validateGateJudge(g);
                    const gate: CanoeSlalomHeatData.gate = {
                        num,
                        direction,
                        judge,
                        fetching: {},
                    }
                    if (rsLocked[i]) {
                        // isLocked
                        gate.isLocked = true;
                    }
                    gates.push(gate);
                });
                dataset.runs[i].gates = gates;
            });
        }

        return dataset;
    }

    export function putData(data: CanoeSlalomHeatData.Data) {
        const sheet = getSheetByHeatName(data.heatName);
        const rowCount = getRowCount(sheet);
        const row = data.runner.row;
        const sheetRow = CONSTS.DATA_TOP_ROW + row;
        if (row >= rowCount) {
            throw new Error(`Invalid row: ${row}`);
        }
        const draftData: CanoeSlalomHeatData.Data = {
            heatName: data.heatName,
            runner: { ...data.runner },
        };
        const draftSystem: CanoeSlalomHeatData.system = { fetching: {} };
        ((draft: CanoeSlalomHeatData.Data) => {
            // 選手データの確認（データロック、楽観的ロック）
            const r = sheet.getRange(sheetRow, CONSTS.RUNNER_COLUMN, 1, CONSTS.RUNNER_LENGTH).getValues()[0];
            const bib = String(r[0]);
            const tag = String(r[1]);
            const isLocked = String(r[2]) != '';
            if ((draft.runner.bib != bib) || (draft.runner.tag != tag)) {
                draft.runner.bib = bib;
                draft.runner.tag = tag;
                draftSystem.fetching.isFailure = true;
            }
            if (isLocked) {
                draftSystem.isLocked = true;
                draftSystem.fetching.isFailure = true;
            }
        })(draftData);
        if (data.started) {
            // スタートタイムの更新
            draftData.started = { ...data.started, ...draftSystem, };
            ((draft: CanoeSlalomHeatData.startedTime) => {
                const range = sheet.getRange(sheetRow, CONSTS.START_COLUMN, 1, CONSTS.START_LENGTH);
                if (draft.fetching.isFailure) {
                    const r = range.getValues()[0];
                    const hms = {
                        hours: r[0],
                        minutes: r[1],
                        seconds: r[2],
                    }
                    draft.seconds = CanoeSlalomHeatData.hmsToSeconds(hms);
                    draft.judge = CanoeSlalomHeatData.validateStartedTimeJudge(r[3]);
                } else {
                    const hms = CanoeSlalomHeatData.secondsToHms(draft.seconds)
                    const r = [
                        hms.hours,
                        hms.minutes,
                        hms.seconds,
                        CanoeSlalomHeatData.validateStartedTimeJudge(draft.judge),
                    ];
                    range.setValues([r]);
                }
            })(draftData.started);
        } else if (data.finished) {
            // ゴールタイムの更新
            draftData.finished = { ...data.finished, ...draftSystem, };
            ((draft: CanoeSlalomHeatData.finishedTime) => {
                const range = sheet.getRange(sheetRow, CONSTS.FINISH_COLUMN, 1, CONSTS.FINISH_LENGTH);
                if (draft.fetching.isFailure) {
                    const r = range.getValues()[0];
                    const hms = {
                        hours: r[0],
                        minutes: r[1],
                        seconds: r[2],
                    }
                    draft.seconds = CanoeSlalomHeatData.hmsToSeconds(hms);
                    draft.judge = CanoeSlalomHeatData.validateFinishedTimeJudge(r[3]);
                } else {
                    const hms = CanoeSlalomHeatData.secondsToHms(draft.seconds)
                    const r = [
                        hms.hours,
                        hms.minutes,
                        hms.seconds,
                        CanoeSlalomHeatData.validateFinishedTimeJudge(draft.judge),
                    ];
                    range.setValues([r]);
                }
            })(draftData.finished);

        } else if (data.gate) {
            // ゲート判定の更新
            draftData.gate = { ...data.gate, ...draftSystem, };
            ((draft: CanoeSlalomHeatData.gate) => {
                const num = draft.num;
                const range = sheet.getRange(sheetRow, num + CONSTS.GATE_COLUMN - 1);
                if (draft.fetching.isFailure) {
                    draft.judge = CanoeSlalomHeatData.validateGateJudge(range.getValue());
                } else {
                    range.setValue(draft.judge);
                }
            })(draftData.gate);
        } else {
            throw new Error('Invalid Data object, must have started, finished or gate.');
        }
        return draftData;
    }

}

export default CanoeSlalomHeatService;
