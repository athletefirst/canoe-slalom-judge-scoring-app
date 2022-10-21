import HtmlTemplateDataSetter from './api/HtmlTemplateDataSetter';
import CanoeSlalomHeatService from './api/CanoeSlalomHeatService';
import WebApiService from './api/WebApiService';
import CanoeSlalomHeatData from './dao/CanoeSlalomHeatData';
import AppConfig from './api/AppConfig';

/**
https://script.google.com/.../exec?h=テストデータ&g=1&n=5

クエリーパラメーター
    h, heatName : ヒート名（シート名）
    g, beginGate : 最初のゲート (=1..30)
    n, gateLength : ゲート数 (=1..)
    s, start : スタート (=1)
    f, finish : ゴール (=1)

 */
function doGet(e: GoogleAppsScript.Events.DoGet) {
    // 認証ページへ遷移してしまうため、path(pathInfo)を設定しないこと。
    const webApp = (appConfig: AppConfig.AppConfig) => {
        const template = HtmlTemplateDataSetter.createTemplateFromFile("index", appConfig);
        return template
            .evaluate()
            .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
            .setTitle("Canoe Slalom Scoring App.");
    };
    const p = e.parameter;
    const appConfig = AppConfig.buildAppConfig(
        p.h ? p.h : p.heatName,
        p.g ? p.g : p.beginGate,
        p.n ? p.n : p.gateLength,
        p.s ? p.s : p.start,
        p.f ? p.f : p.finish
    );
    return webApp(appConfig);
}

/**
curl -i -d '{"operationId":"getHeatsAll", "operationData":{}}' \
  -L 'https://script.google.com/.../exec' \
  -H "Accept: application/json" \
  -H "Content-Type: application/json"
*/
function doPost(e: GoogleAppsScript.Events.DoPost) {
    // 認証ページへ遷移してしまうため、path(pathInfo)を設定しないこと。
    const createJsonResponse = (json: any) => {
        return ContentService
            .createTextOutput(JSON.stringify(json))
            .setMimeType(ContentService.MimeType.JSON);
    }
    const p: WebApiService.PostParameter = JSON.parse(e.postData.contents);
    const result = WebApiService[p.operationId](p.operationData);
    return createJsonResponse({ result });
}

function getDataset(criteria: CanoeSlalomHeatService.Criteria) {
    Logger.log(criteria);
    return CanoeSlalomHeatService.getDataset(criteria);
}

function putData(data: CanoeSlalomHeatData.Data) {
    Logger.log(data);
    return CanoeSlalomHeatService.putData(data);
}
