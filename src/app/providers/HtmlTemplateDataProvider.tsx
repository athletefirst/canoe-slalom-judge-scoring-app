import React, { createContext, useContext } from 'react';
import HtmlTemplateDataSetter from '../../api/HtmlTemplateDataSetter';

const HtmlTemplateDataContext = createContext(HtmlTemplateDataSetter.internal.defaultValue);

export function useData<T = undefined>(defaultData?: T) {
    const { data } = useContext(HtmlTemplateDataContext);
    if (defaultData) {
        return {
            data: data ? data as T : defaultData,
        };
    }
    return {
        data: data as T,
    };
}

export default function HtmlTemplateDataProvider({ children }) {
    const dataContainer = HtmlTemplateDataSetter.internal.getFromData();
    return (
        <HtmlTemplateDataContext.Provider value={dataContainer}>
            {children}
        </HtmlTemplateDataContext.Provider>
    );
}
