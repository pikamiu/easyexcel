package com.sailvan.excel.modelbuild;

import com.sailvan.excel.context.AnalysisContext;
import com.sailvan.excel.event.AnalysisEventListener;
import com.sailvan.excel.exception.ExcelGenerateException;
import com.sailvan.excel.metadata.ExcelHeadProperty;
import com.sailvan.excel.util.TypeUtil;
import net.sf.cglib.beans.BeanMap;

import java.util.List;

/**
 * @author jipengfei
 */
public class ModelBuildEventListener extends AnalysisEventListener {

    @Override
    public void invoke(Object object, AnalysisContext context) {
        if (context.getExcelHeadProperty() != null && context.getExcelHeadProperty().getHeadClazz() != null) {
            try {
                Object resultModel = buildUserModel(context, (List<String>)object);
                context.setCurrentRowAnalysisResult(resultModel);
            } catch (Exception e) {
                throw new ExcelGenerateException(e);
            }
        }
    }

    private Object buildUserModel(AnalysisContext context, List<String> stringList) throws Exception {
        ExcelHeadProperty excelHeadProperty = context.getExcelHeadProperty();
        Object resultModel = excelHeadProperty.getHeadClazz().newInstance();
        if (excelHeadProperty == null) {
            return resultModel;
        }
        BeanMap.create(resultModel).putAll(
            TypeUtil.getFieldValues(stringList, excelHeadProperty, context.use1904WindowDate()));
        return resultModel;
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {

    }
}
