import CanoeSlalomHeatData from '../dao/CanoeSlalomHeatData';
import WebApiData from '../dao/WebApiData';
import CanoeSlalomHeatService from './CanoeSlalomHeatService';

namespace WebApiService {

    export interface PostParameter {
        operationId: string;
        operationData: any;
    }

    export function addHeat(operationData: WebApiData.Parameter) {
        CanoeSlalomHeatService.createNewSheet(operationData.heatName);
        return 'ok';
    }

    export function getHeatsAll(operationData: any) {
        const heats: WebApiData.Heat[] = [];
        const names = CanoeSlalomHeatService.getHeatNameList();
        names.forEach(heatName => {
            heats.push({ heatName });
        });
        return heats;
    }

    export function putGateSettings(operationData: WebApiData.GateSettingsParameter) {
        const heatName = operationData.heatName;
        CanoeSlalomHeatService.putGateSettings(heatName, operationData.gateSettings);
        return 'ok';
    }

    export function putRunners(operationData: WebApiData.RunnersParameter) {
        const heatName = operationData.heatName;
        CanoeSlalomHeatService.putRunners(heatName, operationData.runners);
        return 'ok';
    }

    export function getRecords(operationData: WebApiData.RowsParameter) {
        const records: WebApiData.Record[] = [];
        const criteria: CanoeSlalomHeatService.Criteria = {
            heatName: operationData.heatName,
            started: true,
            finished: true,
            gates: {
                beginGate: 1,
                gateLength: CanoeSlalomHeatData.CONSTS.GATE_MAX,
            },
        };
        const ds = CanoeSlalomHeatService.getDataset(criteria);
        const rowMin = operationData.row1 !== undefined ? operationData.row1 : 0;
        const rowMax = operationData.row2 !== undefined ? operationData.row2 : ds.runs.length - 1;
        ds.runs.forEach(run => {
            if ((run.runner.row < rowMin) || (run.runner.row > rowMax)) {
                return;
            }
            const started: WebApiData.Started = {
                judge: '',
                time: '',
            };
            if (run.started) {
                started.judge = run.started.judge;
                started.time = run.started.seconds === 0 ? '' : String(run.started.seconds);
            }
            const finished: WebApiData.Finished = {
                judge: '',
                time: '',
            };
            if (run.finished) {
                finished.judge = run.finished.judge;
                finished.time = run.finished.seconds === 0 ? '' : String(run.finished.seconds);
            }
            const gates: WebApiData.Gate[] = [];
            if (run.gates) {
                run.gates.forEach(gate => {
                    gates.push({
                        num: gate.num,
                        judge: gate.judge,
                    });
                });
            }
            const record: WebApiData.Record = {
                runner: {
                    row: run.runner.row,
                    bib: run.runner.bib,
                    tag: run.runner.tag,
                    locked: run.runner.locked ? run.runner.locked : '',
                },
                started,
                finished,
                gates,
            };
            records.push(record);
        });
        return records;
    }
}

export default WebApiService;