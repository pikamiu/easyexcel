package com.sailvan.excel.event;

import com.sailvan.excel.context.AnalysisContext;

/**
 * event listener
 * @param <T>
 */
public abstract class AnalysisEventListener<T> {

    /**
     * when analysis one row trigger invoke function
     *
     * @param object  one row data
     * @param context analysis context
     */
    public abstract void invoke(T object, AnalysisContext context);

    /**
     * if have something to do after all  analysis
     *
     * @param context AnalysisContext
     */
    public abstract void doAfterAllAnalysed(AnalysisContext context);
}
