import CanoeSlalomHeatData from '../dao/CanoeSlalomHeatData';
import CanoeSlalomHeatService from './CanoeSlalomHeatService';

const UT_TEST_SHEET_NAME = 'UT-テスト'

function test_CreateNewSheet_CanoeSlalomHeatService() {
    const sheet = CanoeSlalomHeatService.createNewSheet(UT_TEST_SHEET_NAME);
    Logger.log(sheet.getSheetName());
}

function test_GetDataset_CanoeSlalomHeatService() {
    let criteria: CanoeSlalomHeatService.Criteria
    let dataset: CanoeSlalomHeatData.Dataset;

    criteria = {
        heatName: UT_TEST_SHEET_NAME,
        started: true,
        finished: true,
        gates: {
            beginGate: 1,
            gateLength: 30,
        },
    }
    dataset = CanoeSlalomHeatService.getDataset(criteria);
    Logger.log('全て');
    Logger.log(dataset);

    criteria = {
        heatName: UT_TEST_SHEET_NAME,
        started: true,
        finished: true,
    }
    dataset = CanoeSlalomHeatService.getDataset(criteria);
    Logger.log('スタートとゴール');
    Logger.log(dataset);

    criteria = {
        heatName: UT_TEST_SHEET_NAME,
        started: true,
    }
    dataset = CanoeSlalomHeatService.getDataset(criteria);
    Logger.log('スタートのみ');
    Logger.log(dataset);

    criteria = {
        heatName: UT_TEST_SHEET_NAME,
        finished: true,
    }
    dataset = CanoeSlalomHeatService.getDataset(criteria);
    Logger.log('ゴールのみ');
    Logger.log(dataset);

    criteria = {
        heatName: UT_TEST_SHEET_NAME,
        gates: {
            beginGate: 1,
            gateLength: 1,
        },
    }
    dataset = CanoeSlalomHeatService.getDataset(criteria);
    Logger.log('ゲート1');
    Logger.log(dataset);

    criteria = {
        heatName: UT_TEST_SHEET_NAME,
        gates: {
            beginGate: 30,
            gateLength: 1,
        },
    }
    dataset = CanoeSlalomHeatService.getDataset(criteria);
    Logger.log('ゲート30');
    Logger.log(dataset);

    criteria = {
        heatName: UT_TEST_SHEET_NAME,
        gates: {
            beginGate: 13,
            gateLength: 5,
        },
    }
    dataset = CanoeSlalomHeatService.getDataset(criteria);
    Logger.log('ゲート13-17');
    Logger.log(dataset);

}

function test_PutDataSingle_CanoeSlalomHeatService() {
    const criteria: CanoeSlalomHeatService.Criteria = {
        heatName: UT_TEST_SHEET_NAME,
        gates: {
            beginGate: 1,
            gateLength: 30,
        },
    }
    const dataset = CanoeSlalomHeatService.getDataset(criteria);

    let run: CanoeSlalomHeatData.run;
    let runner: CanoeSlalomHeatData.runner;
    let gate: CanoeSlalomHeatData.gate;

    let modGate: CanoeSlalomHeatData.gate;
    let modRunner: CanoeSlalomHeatData.runner;

    let modData:CanoeSlalomHeatData.Data;
    let result: any;

    run = dataset.runs[0];
    runner = run.runner;
    if (run.gates) {
        gate = run.gates[29];

        modGate = {
            num: gate.num,
            judge: '0',
            fetching: {},
        }

        modRunner = {
            row: runner.row,
            bib: runner.bib,
            tag: runner.tag,
        }

        modData = {
            heatName: dataset.heatName,
            runner: modRunner,
            gate: modGate,
        };
        result = CanoeSlalomHeatService.putData(modData);
        Logger.log(result);
    }
}

function test_PutData_CanoeSlalomHeatService() {
    const heatName = UT_TEST_SHEET_NAME;
    const criteria: CanoeSlalomHeatService.Criteria = {
        heatName,
        gates: {
            beginGate: 1,
            gateLength: 30,
        },
    }
    const dataset = CanoeSlalomHeatService.getDataset(criteria);

    const putGate = (runNumber: number, gateNumber: number, newJudge: any) => {
        const runIndex = runNumber - 1;
        const gateIndex = gateNumber - 1;
        const run = dataset.runs[runIndex];
        const runner = run.runner;
        if (run.gates) {
            const gate: CanoeSlalomHeatData.gate = {
                ...run.gates[gateIndex],
                judge: CanoeSlalomHeatData.validateGateJudge(newJudge),
                fetching: {}
            };
            const data: CanoeSlalomHeatData.Data = {
                heatName,
                runner,
                gate,
            }
            const result = CanoeSlalomHeatService.putData(data);
            return result;
        }
    }
    let result: any;
    result = putGate(1, 1, '0');
    Logger.log(result);
    result = putGate(1, 30, '2');
    Logger.log(result);
    result = putGate(3, 1, '50');
    Logger.log(result);
    result = putGate(3, 30, 'DSQ');
    Logger.log(result);

    result = putGate(2, 10, '');
    Logger.log(result);

}
