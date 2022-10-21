namespace AppConfig {

    export type AppConfig = {
        heatName: string;
        beginGate: number;
        gateLength: number;
        start: number;
        finish: number;
    }

    export const defaultValue: AppConfig = {
        heatName: '',
        beginGate: 0,
        gateLength: 0,
        start: 0,
        finish: 0,
    };

    export function buildAppConfig(heatName: any, beginGate: any, gateLength: any, start: any, finish: any): AppConfig {
        return {
            heatName: String(heatName),
            beginGate: Number(beginGate),
            gateLength: Number(gateLength),
            start: Number(start),
            finish: Number(finish),
        }
    }

}

export default AppConfig;