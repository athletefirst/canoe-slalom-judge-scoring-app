import CanoeSlalomHeatData from './CanoeSlalomHeatData';

namespace WebApiData {

    export interface Parameter {
        heatName: string;
    }

    export interface RowsParameter extends Parameter {
        row1?: number;
        row2?: number;
    }

    export interface GatesParameter extends Parameter {
        num1?: number;
        num2?: number;
    }

    export interface GateSettingsParameter extends Parameter {
        gateSettings: GateSetting[];
    }

    export interface RunnersParameter extends Parameter {
        runners: Runner[];
    }

    export type Heat = {
        heatName: string;
    };

    export type GateSetting = {
        num: number;
        direction: CanoeSlalomHeatData.gateType;
    }

    export type Record = {
        runner: Runner;
        started: Started;
        finished: Finished;
        gates: Gate[];
    };

    export type Runner = {
        row: number;
        bib: string;
        tag: string;
        locked: string;
    };

    export type Started = {
        judge: string;
        time: string;
    };

    export type Finished = {
        judge: string;
        time: string;
    };

    export type Gate = {
        num: number;
        judge: string;
    };
}

export default WebApiData;
