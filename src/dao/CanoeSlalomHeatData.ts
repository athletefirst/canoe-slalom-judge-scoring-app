namespace CanoeSlalomHeatData {

    export namespace CONSTS {

        /**
         * 判定 - 未
         */
        export const JUDGE_NONE = '';
        export type judgeNone = '';

        /**
         * 判定 - 棄権 (スタートタイム) 
         */
        export const JUDGE_DNS = 'DNS';
        export type judgeDns = 'DNS';

        /**
         * 判定 - スタート済 (スタートタイム) 
         */
        export const JUDGE_STARTED = 'STARTED';
        export type judgeStarted = 'STARTED';

        /**
         * 判定 - 途中棄権 (ゴールタイム) 
         */
        export const JUDGE_DNF = 'DNF';
        export type judgeDnf = 'DNF';

        /**
         * 判定 - ゴール済 (ゴールタイム) 
         */
        export const JUDGE_FINISHED = 'FINISHED';
        export type judgeFinished = 'FINISHED';

        /**
         * 判定 - ゴール済 (ゴールタイム) - 50ペナルティー付き(チームレース15秒ルール)
         */
        export const JUDGE_FINISHED_50 = 'FINISHED_50';
        export type judgeFinished50 = 'FINISHED_50';

        /**
         * 判定 - ペナルティ0（ゲート判定）
         */
        export const JUDGE_P0 = '0';
        export type judgeP0 = '0';

        /**
         * 判定 - ペナルティ2（ゲート判定）
         */
        export const JUDGE_P2 = '2';
        export type judgeP2 = '2';

        /**
         * 判定 - ペナルティ50（ゲート判定）
         */
        export const JUDGE_P50 = '50';
        export type judgeP50 = '50';

        /**
         * 判定 - 失格（スタートタイム/ゲート判定/ゴールタイム）
         */
        export const JUDGE_DSQ = 'DSQ';
        export type judgeDsq = 'DSQ';

        /**
         * 最大ゲート番号
         */
        export const GATE_MAX = 30;

        /**
         * ダウンストリームゲート
         * （UP、FREEの何れでもなけらば、DOWN）
         */
        export const GATE_DOWNSTREAM = 'DOWN';
        export type gateDownstream = 'DOWN';

        /**
         * アップストリームゲート
         */
        export const GATE_UPSTREAM = 'UP';
        export type gateUpstream = 'UP';

        /**
         * 廃止ゲート(ゲート番号設置後に廃止された等)
         */
        export const GATE_FREE = 'FREE';
        export type gateFree = 'FREE';

    }

    /**
     * 単体データ
     * (started, finished, gate) の何れか１つのみが必要
     */
    export type Data = {
        /**
         * ヒート名（シート名）
         */
        heatName: string;
        /**
         * 選手データ
         */
        runner: runner;
        /**
         * スタートタイム
         */
        started?: startedTime;
        /**
         * フィニッシュタイム
         */
        finished?: finishedTime;
        /**
         * ゲート判定
         */
        gate?: gate;
    };

    /**
     * 選手・タイム・ペナルティのデータセット
     */
    export type Dataset = {
        /**
         * スプレッドシートのヒート名
         */
        heatName: string;
        /**
         * 選手・タイム・ペナルティのレコードセット
         */
        runs: run[];
    }

    /**
     * 選手・タイム・ペナルティ
     */
    export type run = {
        /**
         * 選手データ
         */
        runner: runner,
        /**
         * スタートタイム
         */
        started?: startedTime;
        /**
         * ゴールタイム
         */
        finished?: finishedTime;
        /**
         * ゲート判定(要素数:1以上)
         */
        gates?: gate[];
    }

    /**
     * 選手データ
     */
    export type runner = {
        /**
         * [PK] シートの行番号
         */
        row: number;
        /**
         * ゼッケン番号
         */
        bib: string;
        /**
         * 識別情報 - ヒート、カテゴリ、チーム名、選手名等
         */
        tag: string;
        /**
         * 行ロック情報
         */
        locked?: string;
    }

    /**
     * スタートタイム
     */
    export type startedTime = {
        /**
         * スタートタイム判定
         */
        judge: startedTimeJudge;
    } & time

    type startedTimeJudge = CONSTS.judgeNone | CONSTS.judgeDns | CONSTS.judgeStarted | CONSTS.judgeDsq;

    /**
     * スタートタイム判定の検証と変換
     * @param value 検証する値
     * @returns 変換された値（無効な値の場合、例外発生）
     */
    export function validateStartedTimeJudge(value: any): startedTimeJudge {
        switch (String(value).toUpperCase()) {
            case CONSTS.JUDGE_NONE:
                return CONSTS.JUDGE_NONE;
                break;
            case CONSTS.JUDGE_DNS:
                return CONSTS.JUDGE_DNS;
                break;
            case CONSTS.JUDGE_STARTED:
                return CONSTS.JUDGE_STARTED;
                break;
            case CONSTS.JUDGE_DSQ:
                return CONSTS.JUDGE_DSQ;
                break;
            default:
                throw new Error(`Invalid StartedTime judge: ${value}`);
                break;
        }
    }

    /**
     * ゴールタイム
     */
    export type finishedTime = {
        /**
         * ゴールタイム判定
         */
        judge: finishedTimeJudge;
    } & time

    type finishedTimeJudge = CONSTS.judgeNone | CONSTS.judgeDnf | CONSTS.judgeFinished | CONSTS.judgeFinished50 | CONSTS.judgeDsq;

    /**
     * ゴールタイム判定の検証と変換
     * @param value 検証する値
     * @returns 変換された値（無効な値の場合、例外発生）
     */
    export function validateFinishedTimeJudge(value: any): finishedTimeJudge {
        switch (String(value).toUpperCase()) {
            case CONSTS.JUDGE_NONE:
                return CONSTS.JUDGE_NONE;
                break;
            case CONSTS.JUDGE_DNF:
                return CONSTS.JUDGE_DNF;
                break;
            case CONSTS.JUDGE_FINISHED:
                return CONSTS.JUDGE_FINISHED;
                break;
            case CONSTS.JUDGE_FINISHED_50:
                return CONSTS.JUDGE_FINISHED_50;
                break;
            case CONSTS.JUDGE_DSQ:
                return CONSTS.JUDGE_DSQ;
                break;
            default:
                throw new Error(`Invalid FinishedTime judge: ${value}`);
                break;
        }
    }

    /**
     * スタートタイム・ゴールタイム
     */
    type time = {
        /**
         * 秒(小数点以下3桁まで)
         */
        seconds: number;
    } & system

    /**
     * 時・分・秒（小数点以下3桁）
     */
    type hms = {
        hours: number,
        minutes: number,
        seconds: number,
    }

    /**
     * タイム変換
     * @param hms 時・分・秒
     * @returns 秒
     */
    export function hmsToSeconds(hms: hms) {
        return (
            (
                Number(hms.hours) * 60  // 分
                + Number(hms.minutes)
            ) * 60                      // 秒
            + Number(hms.seconds)
        );
    }

    /**
     * タイム変換
     * @param seconds 秒 
     * @returns 時・分・秒
     */
    export function secondsToHms(seconds: number): hms {
        const hours = Math.floor(seconds / (60 * 60));
        const minutes = Math.floor((seconds - hours * 60 * 60) / 60);
        return {
            hours,
            minutes,
            seconds: (seconds * 1000 - (hours * 60 + minutes) * 60 * 1000) / 1000,
        };
    }

    /**
     * ゲート設定
     */
    export type gateSetting = {
        /**
         * <<PK>> ゲート番号(1～30)
         */
        num: number;
        /**
         * 進行方向
         */
        direction?: gateType;
    }

    /**
     * ゲート判定
     */
    export type gate = {
        /**
         * <<PK>> ゲート番号(1～30)
         */
        num: number;
        /**
         * 進行方向
         */
        direction?: gateType;
        /**
         * ゲート判定
         */
        judge: gateJudge;
    } & system;
    export type gateType = CONSTS.gateDownstream | CONSTS.gateUpstream | CONSTS.gateFree;
    type gateJudge = CONSTS.judgeNone | CONSTS.judgeP0 | CONSTS.judgeP2 | CONSTS.judgeP50 | CONSTS.judgeDsq;

    /**
     * ゲート番号の検証
     * @param value 検証するゲート番号
     * @returns 検証済みのゲート番号（無効な場合、例外発生）
     */
    export function validateGateNum(value: any): number {
        const num = Number(value);
        if ((num < 1) || (num > CONSTS.GATE_MAX)) {
            throw new Error(`Invalid Gate num: ${value} [1-${CONSTS.GATE_MAX}]`);
        }
        return num;
    }

    /**
     * ゲートタイプへの変換
     * @param values 変換する値のリスト
     * @returns ゲートタイプのリスト
     */
    export function toStringGateTypeList(gateTypes: gateType[]) {
        const ret: any[] = [];
        const toStringGateType = (value: gateType) => {
            switch (value) {
                case CONSTS.GATE_FREE:
                    return CONSTS.GATE_FREE;
                    break;
                case CONSTS.GATE_UPSTREAM:
                    return CONSTS.GATE_UPSTREAM;
                    break;
                default:
                    return '';
                    break;
            }
        };
        gateTypes.forEach(v => {
            ret.push(toStringGateType(v));
        });
        return ret;
    }

    /**
     * ゲートタイプへの変換
     * @param values 変換する値のリスト
     * @returns ゲートタイプのリスト
     */
    export function convGateTypeList(values: any[]) {
        const gateTypes: gateType[] = [];
        values.forEach(v => {
            gateTypes.push(convGateType(v));
        });
        return gateTypes;
    }

    /**
     * ゲートタイプへの変換
     * @param value 変換する値
     * @returns ゲートタイプ
     */
    export function convGateType(value: any) {
        switch (String(value).toUpperCase()) {
            case CONSTS.GATE_FREE:
                return CONSTS.GATE_FREE;
                break;
            case CONSTS.GATE_UPSTREAM:
                return CONSTS.GATE_UPSTREAM;
                break;
            default:
                return CONSTS.GATE_DOWNSTREAM;
                break;
        }
    };
    /**
     * ゲート判定の検証と変換
     * @param value 検証する値
     * @returns 変換された値（無効な値の場合、例外発生）
     */
    export function validateGateJudge(value: any): gateJudge {
        switch (String(value).toUpperCase()) {
            case CONSTS.JUDGE_NONE:
                return CONSTS.JUDGE_NONE;
                break;
            case CONSTS.JUDGE_P0:
                return CONSTS.JUDGE_P0;
                break;
            case CONSTS.JUDGE_P2:
                return CONSTS.JUDGE_P2;
                break;
            case CONSTS.JUDGE_P50:
                return CONSTS.JUDGE_P50;
                break;
            case CONSTS.JUDGE_DSQ:
                return CONSTS.JUDGE_DSQ;
                break;
            default:
                throw new Error(`Invalid Gate judge: ${value}`);
                break;
        }
    }

    /**
     * システム項目
     */
    export type system = {
        /**
         * データロック済みフラグ
         */
        isLocked?: LOCKED;
        /**
         * データ反映状況(何もなければ、成功)
         */
        fetching: fetching;
    }

    /**
     * データロック済み状態
     */
    type LOCKED = true;

    /**
     * データ反映状況(何もなければ、成功)
     */
    type fetching = {
        /**
         * データ通信中
         */
        isLoading?: IS_LOADING;
        /**
         * 処理エラーあり
         * （更新失敗も同時にtrueとなる）
         */
        hasError?: HAS_ERROR;
        /**
         * 更新失敗
         */
        isFailure?: IS_FAILURE;
    }

    /**
     * データ通信中状態
     */
    type IS_LOADING = true;

    /**
     * 処理エラーあり
     */
    type HAS_ERROR = true;

    /**
     * 更新失敗（処理エラーを含む）
     */
    type IS_FAILURE = true;

}

export default CanoeSlalomHeatData;
