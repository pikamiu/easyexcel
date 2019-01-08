package com.sailvan.excel.write;

import com.sailvan.excel.context.WriteContext;
import com.sailvan.excel.event.WriteHandler;
import com.sailvan.excel.exception.ExcelGenerateException;
import com.sailvan.excel.metadata.ExcelColumnProperty;
import com.sailvan.excel.metadata.Sheet;
import com.sailvan.excel.metadata.Table;
import com.sailvan.excel.support.ExcelTypeEnum;
import com.sailvan.excel.util.CollectionUtils;
import com.sailvan.excel.util.JSONUtil;
import com.sailvan.excel.util.ListUtil;
import com.sailvan.excel.util.POITempFile;
import com.sailvan.excel.util.StyleUtil;
import com.sailvan.excel.util.TypeUtil;
import com.sailvan.excel.util.WorkBookUtil;
import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;

import net.sf.cglib.beans.BeanMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;


/**
 * @author jipengfei
 */
public class ExcelBuilderImpl implements ExcelBuilder {

    private WriteContext context;

    public ExcelBuilderImpl(InputStream templateInputStream,
                            OutputStream out,
                            ExcelTypeEnum excelType,
                            boolean needHead, WriteHandler writeHandler) {
        try {
            //初始化时候创建临时缓存目录，用于规避POI在并发写bug
            POITempFile.createPOIFilesDirectory();
            context = new WriteContext(templateInputStream, out, excelType, needHead, writeHandler);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public void addContent(List data, int startRow) {
        if (CollectionUtils.isEmpty(data)) {
            return;
        }
        int rowNum = context.getCurrentSheet().getLastRowNum();
        if (rowNum == 0) {
            Row row = context.getCurrentSheet().getRow(0);
            if (row == null) {
                if (context.getExcelHeadProperty() == null || !context.needHead()) {
                    rowNum = -1;
                }
            }
        }
        if (rowNum < startRow) {
            rowNum = startRow;
        }
        for (int i = 0; i < data.size(); i++) {
            int n = i + rowNum + 1;
            addOneRowOfDataToExcel(data.get(i), n);
        }
    }

    @Override
    public void addContent(List data, Sheet sheetParam) {
        context.currentSheet(sheetParam);
        addContent(data, sheetParam.getStartRow());
    }

    @Override
    public void addContent(List data, Sheet sheetParam, Table table) {
        context.currentSheet(sheetParam);
        context.currentTable(table);
        addContent(data, sheetParam.getStartRow());
    }

    @Override
    public void addContent(Map data, int startRow) {
        if (CollectionUtils.isEmpty(data)) {
            return;
        }
        int rowNum = context.getCurrentSheet().getLastRowNum();
        if (rowNum == 0) {
            Row row = context.getCurrentSheet().getRow(0);
            if (row == null) {
                if (context.getExcelHeadProperty() == null || !context.needHead()) {
                    rowNum = -1;
                }
            }
        }
        if (rowNum < startRow) {
            rowNum = startRow;
        }
        for (int i = 0; i < data.size(); i++) {
            int n = i + rowNum + 1;
            addOneRowOfDataToExcel(data.get(i), n);
        }
    }

    @Override
    public void addContent(Map data, Sheet sheetParam) {
        context.currentSheet(sheetParam);
        addContent(data, sheetParam.getStartRow());
    }

    @Override
    public void addHead(Object object, Sheet sheetParam) {
        if (object instanceof String) {
            addJsonHead((String) object, sheetParam);
        }
    }

    @Override
    public void merge(int firstRow, int lastRow, int firstCol, int lastCol) {
        CellRangeAddress cra = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        context.getCurrentSheet().addMergedRegion(cra);
    }

    @Override
    public void finish() {
        try {
            context.getWorkbook().write(context.getOutputStream());
            context.getWorkbook().close();
        } catch (IOException e) {
            throw new ExcelGenerateException("IO error", e);
        }
    }

    private void addBasicTypeToExcel(List<Object> oneRowData, Row row) {
        if (CollectionUtils.isEmpty(oneRowData)) {
            return;
        }
        for (int i = 0; i < oneRowData.size(); i++) {
            Object cellValue = oneRowData.get(i);
            Cell cell = WorkBookUtil.createCell(row, i, context.getCurrentContentStyle(), cellValue,
                TypeUtil.isNum(cellValue));
            StyleUtil.autoColumnSize(context.getCurrentSheet(), context.getCurrentSheetParam(), i);
            if (null != context.getAfterWriteHandler()) {
                context.getAfterWriteHandler().cell(i, cell, cellValue);
            }
        }
    }

    private void addJsonHead(String str, Sheet sheetParam) {
        if (sheetParam.getHead() == null && sheetParam.getSingleHead() == null) {
            List<String> l = JSONUtil.jsonParse(str);
            List<List<String>> jsonHead = ListUtil.listWarp(JSONUtil.jsonParse(str));
            sheetParam.setHead(jsonHead);
            sheetParam.setContentTitle(l);
        }
    }

    private void addJsonToExcel(Object oneRowData, Row row) {
        if (oneRowData == null) {
            return;
        }
        List<String> heads = context.getContentTitle();
        for (int i = 0; i < heads.size(); i++) {
            JSONObject json = JSON.parseObject(String.valueOf(oneRowData));
            Object cellValue = json.get(heads.get(i));
            Cell cell = WorkBookUtil.createCell(row, i, context.getCurrentContentStyle(), cellValue,
                    TypeUtil.isNum(cellValue));
            StyleUtil.autoColumnSize(context.getCurrentSheet(), context.getCurrentSheetParam(), i);
            if (null != context.getAfterWriteHandler()) {
                context.getAfterWriteHandler().cell(i, cell, cellValue);
            }
        }
    }

    private void addMapToExcel(Object oneRowData, Row row) {
        if (oneRowData == null) {
            return;
        }
        List<String> contentTitle = context.getContentTitle();
        if (contentTitle == null) {
            throw new ExcelGenerateException("contentTitle can't be empty");
        }
        for (int i = 0; i < contentTitle.size(); i++) {
            Map<String, Object> map = (Map<String, Object>) oneRowData;
            Object cellValue = map.get(contentTitle.get(i));
            Cell cell = WorkBookUtil.createCell(row, i, context.getCurrentContentStyle(), cellValue,
                    TypeUtil.isNum(cellValue));
            StyleUtil.autoColumnSize(context.getCurrentSheet(), context.getCurrentSheetParam(), i);
            if (null != context.getAfterWriteHandler()) {
                context.getAfterWriteHandler().cell(i, cell, cellValue);
            }
        }
    }

    private void addJavaObjectToExcel(Object oneRowData, Row row) {
        int i = 0;
        BeanMap beanMap = BeanMap.create(oneRowData);
        for (ExcelColumnProperty excelHeadProperty : context.getExcelHeadProperty().getColumnPropertyList()) {
            String cellValue = TypeUtil.getFieldStringValue(beanMap, excelHeadProperty.getField().getName(),
                excelHeadProperty.getFormat());
            /*
              // give up each column style set to compatible framework annotations
              BaseRowModel baseRowModel = (BaseRowModel)oneRowData;
              CellStyle cellStyle = baseRowModel.getStyle(i) != null ? baseRowModel.getStyle(i)
                  : context.getCurrentContentStyle();
            */

            CellStyle cellStyle = context.getCurrentContentStyle();
            Cell cell = WorkBookUtil.createCell(row, i, cellStyle, cellValue,
                TypeUtil.isNum(excelHeadProperty.getField()));

            // set auto size columns
            StyleUtil.autoColumnSize(context.getCurrentSheet(), context.getCurrentSheetParam(), i);

            if (null != context.getAfterWriteHandler()) {
                context.getAfterWriteHandler().cell(i, cell, cellValue);
            }
            i++;
        }

    }

    private void addOneRowOfDataToExcel(Object oneRowData, int n) {
        Row row = WorkBookUtil.createRow(context.getCurrentSheet(), n);
        if (null != context.getAfterWriteHandler()) {
            context.getAfterWriteHandler().row(n, row);
        }
        if (oneRowData instanceof List) {
            addBasicTypeToExcel((List) oneRowData, row);
        } else if (oneRowData instanceof String) {
            addJsonToExcel(oneRowData, row);
        } else if(oneRowData instanceof Map) {
            addMapToExcel(oneRowData, row);
        } else {
            addJavaObjectToExcel(oneRowData, row);
        }
    }
}
