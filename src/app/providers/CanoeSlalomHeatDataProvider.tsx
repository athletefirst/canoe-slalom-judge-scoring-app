import React, { createContext, useContext, useState, useEffect } from 'react';
import { useImmerReducer } from 'use-immer';
import { GASClient } from 'gas-client';
const { serverFunctions } = new GASClient();
import CanoeSlalomHeatService from '../../api/CanoeSlalomHeatService';
import CanoeSlalomHeatData from '../../dao/CanoeSlalomHeatData';
import AppConfig from '../../api/AppConfig';
import { useData as usePushedData } from './HtmlTemplateDataProvider';

// ToDo: immer
// https://immerjs.github.io/immer/example-setstate#useimmerreducer

type action = actionDatasetLoaded | actionDataChanged | actionDataUpdated | actionErrorUpdating;

/**
 * データセット取得完了
 */
const ACTION_TYPE_LOADED = 'loaded';
type actionTypeLoaded = 'loaded';

type actionDatasetLoaded = {
    type: actionTypeLoaded;
    payload: CanoeSlalomHeatData.Dataset;
};

/**
 * 画面データ変更
 */
const ACTION_TYPE_CHANGED = 'changed';
type actionTypeChanged = 'changed';

type actionDataChanged = {
    type: actionTypeChanged;
    payload: CanoeSlalomHeatData.Data;
}

/**
 * データ反映完了（成功、失敗、ロック済）
 */
const ACTION_TYPE_UPDATED = 'updated';
type actionTypeUpdated = 'updated';

type actionDataUpdated = {
    type: actionTypeUpdated;
    payload: CanoeSlalomHeatData.Data;
}

/**
 * データ反映エラー
 */
const ACTION_TYPE_ERROR_UPDATING = 'error_updating';
type actionTypeErrorUpdating = 'error_updating';

type actionErrorUpdating = {
    type: actionTypeErrorUpdating;
    payload: CanoeSlalomHeatData.Data;
}

function reducer(draft: CanoeSlalomHeatData.Dataset, action: action) {
    switch (action.type) {
        case ACTION_TYPE_LOADED:
            // ロード完了 - データセットの取得完了
            ((action: actionDatasetLoaded) => {
                // draftへ全てコピー
                draft.heatName = action.payload.heatName;
                draft.runs = action.payload.runs;
            })(action);
            break;
        case ACTION_TYPE_CHANGED:
            // 変更
            ((action: actionDataChanged) => {
                const row = action.payload.runner.row;
                const run = draft.runs.find(run => run.runner.row === row);
                if (run) {
                    if (action.payload.started) {
                        // スタートタイムの変更
                        if (run.started) {
                            run.started.seconds = action.payload.started.seconds;
                            run.started.judge = action.payload.started.judge;
                            run.started.fetching.isLoading = true;
                        }
                    } else if (action.payload.finished) {
                        // ゴールタイムの変更
                        if (run.finished) {
                            run.finished.seconds = action.payload.finished.seconds;
                            run.finished.judge = action.payload.finished.judge;
                            run.finished.fetching.isLoading = true;
                        }
                    } else if (action.payload.gate) {
                        // ゲート判定の変更
                        if (run.gates) {
                            const num = action.payload.gate.num;
                            const gate = run.gates.find(gate => gate.num === num);
                            if (gate) {
                                gate.judge = action.payload.gate.judge;
                                gate.fetching.isLoading = true;
                            }
                        }
                    }
                }
            })(action);
            break;
        case ACTION_TYPE_UPDATED:
            // 更新 - サーバー反映処理完了（エラーを含む）
            ((action: actionDataUpdated) => {
                const row = action.payload.runner.row;
                const run = draft.runs.find(run => run.runner.row === row);
                if (run) {
                    let isLocked = false;
                    if (action.payload.started) {
                        // スタートタイムの反映
                        if (run.started) {
                            isLocked = action.payload.started.isLocked === true;
                            if (isLocked) {
                                run.started.isLocked = true;
                                run.started.seconds = action.payload.started.seconds;
                                run.started.judge = action.payload.started.judge;
                            }
                            run.started.fetching = { ...action.payload.started.fetching }
                        }
                    } else if (action.payload.finished) {
                        // ゴールタイムの反映
                        if (run.finished) {
                            isLocked = action.payload.finished.isLocked === true;
                            if (isLocked) {
                                run.finished.isLocked = true;
                                run.finished.seconds = action.payload.finished.seconds;
                                run.finished.judge = action.payload.finished.judge;
                            }
                            run.finished.fetching = { ...action.payload.finished.fetching }
                        }
                    } else if (action.payload.gate) {
                        // ゲート判定の反映
                        if (run.gates) {
                            isLocked = action.payload.gate.isLocked === true;
                            const num = action.payload.gate.num;
                            const gate = run.gates.find(gate => gate.num === num);
                            if (gate) {
                                if (isLocked) {
                                    gate.isLocked = true;
                                    gate.judge = action.payload.gate.judge;
                                }
                                gate.fetching = { ...action.payload.gate.fetching }
                            }
                        }
                    }
                    if (isLocked) {
                        run.runner.bib = action.payload.runner.bib;
                        run.runner.tag = action.payload.runner.tag;
                    }
                }
            })(action);
            break;
        case ACTION_TYPE_ERROR_UPDATING:
            // データ反映エラー
            ((action: actionErrorUpdating) => {
                const row = action.payload.runner.row;
                const run = draft.runs.find(run => run.runner.row === row);
                if (run) {
                    if (action.payload.started) {
                        // スタートタイムの反映エラー
                        if (run.started) {
                            run.started.fetching = {
                                hasError: true,
                            };
                        }
                    } else if (action.payload.finished) {
                        // ゴールタイムの反映エラー
                        if (run.finished) {
                            run.finished.fetching = {
                                hasError: true,
                            };
                        }
                    } else if (action.payload.gate) {
                        // ゲート判定の反映エラー
                        if (run.gates) {
                            const num = action.payload.gate.num;
                            const gate = run.gates.find(gate => gate.num === num);
                            if (gate) {
                                gate.fetching = {
                                    hasError: true,
                                };
                            }
                        }
                    }
                }
            })(action);
            break;
        default:
            break;
    }
}

const useDataset = (initialDataset: CanoeSlalomHeatData.Dataset, criteria: CanoeSlalomHeatService.Criteria): CanoeSlalomHeatDataContextType => {
    const [dataset, dispatch] = useImmerReducer(reducer, initialDataset);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState();
    const loadDataset = () => {
        if (loading) {
            return;
        }
        setLoading(true);
        setError(undefined);
        serverFunctions.getDataset(criteria)
            .then(dispatchLoaded)
            .catch(setError)
            .finally(() => setLoading(false));
    }
    const getDataAttrs = (row: number) => {
        const heatName = dataset.heatName;
        let bib = '';
        let tag = '';
        const runs = dataset.runs.filter(run => run.runner.row == row);
        if (runs.length > 0) {
            bib = runs[0].runner.bib;
            tag = runs[0].runner.tag;
        }
        return { heatName, bib, tag, fetching: {}, }
    }
    const dispatchUpdated = (data: CanoeSlalomHeatData.Data) => {
        dispatch({
            type: ACTION_TYPE_UPDATED,
            payload: data,
        });
    }
    const dispatchErrorUpdating = (data: CanoeSlalomHeatData.Data, error: any) => {
        dispatch({
            type: ACTION_TYPE_ERROR_UPDATING,
            payload: data,
        });
    }
    const dispatchChanged = (data: CanoeSlalomHeatData.Data) => {
        dispatch({
            type: ACTION_TYPE_CHANGED,
            payload: data,
        });
        serverFunctions.putData(data)
            .then(dispatchUpdated)
            .catch(e => dispatchErrorUpdating(data, e));
    }
    const dispatchLoaded = (dataset: CanoeSlalomHeatData.Dataset) => {
        dispatch({
            type: ACTION_TYPE_LOADED,
            payload: dataset,
        });
    };
    const setStartedTime = (row: number, seconds: number, judge: any) => {
        const { heatName, bib, tag, fetching } = getDataAttrs(row);
        const data: CanoeSlalomHeatData.Data = {
            heatName,
            runner: {
                row,
                bib,
                tag,
            },
            started: {
                seconds,
                judge,
                fetching,
            }
        }
        dispatchChanged(data);
    }
    const setFinishedTime = (row: number, seconds: number, judge: any) => {
        const { heatName, bib, tag, fetching } = getDataAttrs(row);
        const data: CanoeSlalomHeatData.Data = {
            heatName,
            runner: {
                row,
                bib,
                tag,
            },
            finished: {
                seconds,
                judge,
                fetching,
            }
        }
        dispatchChanged(data);
    }
    const setGateJudge = (row: number, num: number, judge: any) => {
        const { heatName, bib, tag, fetching } = getDataAttrs(row);
        const data: CanoeSlalomHeatData.Data = {
            heatName,
            runner: {
                row,
                bib,
                tag,
            },
            gate: {
                num,
                judge,
                fetching,
            }
        }
        dispatchChanged(data);
    }
    useEffect(() => {
        loadDataset();
    }, []);
    return {
        dataset, loading, error, setStartedTime, setFinishedTime, setGateJudge, loadDataset
    }
}

type CanoeSlalomHeatDataContextType = {
    /**
     * データセット
     */
    dataset: CanoeSlalomHeatData.Dataset;
    /**
     * データセット取得中（起動時）
     */
    loading: boolean;
    /**
     * データセット取得エラー（起動時）
     */
    error?: any;
    /**
     * スタートタイム変更設定（アクション）
     */
    setStartedTime: (row: number, seconds: number, judge: any) => void;
    /**
     * ゴールタイム変更設定（アクション）
     */
    setFinishedTime: (row: number, seconds: number, judge: any) => void;
    /**
     * ゲート判定変更設定（アクション）
     */
    setGateJudge: (row: number, num: number, judge: any) => void;
    /**
     * Datasetの読み込み
     */
    loadDataset: () => void;
};

const defaultValue: CanoeSlalomHeatDataContextType = {
    dataset: {
        heatName: '',
        runs: [],
    },
    loading: false,
    setStartedTime: (row, seconds, judge) => undefined,
    setFinishedTime: (row, seconds, judge) => undefined,
    setGateJudge: (row, num, judge) => undefined,
    loadDataset: () => undefined,
}

const CanoeSlalomHeatDataContext = createContext(defaultValue);

export const useData = () => useContext(CanoeSlalomHeatDataContext);

export default function CanoeSlalomHeatDataProvider({ children }) {
    const { data } = usePushedData<AppConfig.AppConfig>();
    const criteria: CanoeSlalomHeatService.Criteria = {
        heatName: data.heatName,
    }
    if (data.gateLength > 0) {
        criteria.gates = {
            beginGate: data.beginGate,
            gateLength: data.gateLength,
        }
    }
    if (data.start > 0) {
        criteria.started = true;
    }
    if (data.finish > 0) {
        criteria.finished = true;
    }
    const providerValue = useDataset({ ...defaultValue.dataset, heatName: data.heatName, }, criteria)
    return (
        <CanoeSlalomHeatDataContext.Provider value={providerValue}>
            {children}
        </CanoeSlalomHeatDataContext.Provider>
    );
}
