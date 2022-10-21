namespace HtmlTemplateDataSetter {

    /**
     * Creates a new HtmlTemplate object from a file in the code editor,
     * using HtmlTemplateDataProvider component.
     * @param filename the name of the file to use.
     * @param data the data pushing to template.
     * @returns the new HtmlTemplate object with setData() method added.
     */
    export function createTemplateFromFile(filename: string, data?: any): HtmlTemplate {
        const template = HtmlService.createTemplateFromFile(filename);
        template.setData = (data: any) => internal.pushToData(template, data);
        return data ? template.setData(data) : template;
    }

    interface HtmlTemplate extends GoogleAppsScript.HTML.HtmlTemplate {
        setData(data: any): HtmlTemplate;
    }

    export namespace internal {

        /**
         * Pushing Variable: `data`
         * https://developers.google.com/apps-script/guides/html/templates#pushing_variables_to_templates     
         */

        interface DataContainer {
            data: any;
        };

        const htmlTemplateDataContainerString = '<? try {?><?= data ?><? } catch {} ?>';

        export function pushToData(template: GoogleAppsScript.HTML.HtmlTemplate, data: any) {
            const container: DataContainer = { data };
            template.data = JSON.stringify(container);
            return template;
        }

        export function getFromData() {
            let container: DataContainer;
            try {
                container = JSON.parse(htmlTemplateDataContainerString);
            } catch (error) {
                container = { data: null, }
            }
            return container;
        }

        export const defaultValue: DataContainer = {
            data: null,
        }

    }

}

export default HtmlTemplateDataSetter;
