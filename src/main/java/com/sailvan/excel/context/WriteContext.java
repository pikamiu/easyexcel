package com.sailvan.excel.context;

import com.sailvan.excel.event.WriteHandler;
import com.sailvan.excel.metadata.BaseRowModel;
import com.sailvan.excel.metadata.ExcelHeadProperty;
import com.sailvan.excel.support.ExcelTypeEnum;
import com.sailvan.excel.util.CollectionUtils;
import com.sailvan.excel.util.ListUtil;
import com.sailvan.excel.util.StyleUtil;
import com.sailvan.excel.util.WorkBookUtil;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

import static com.sailvan.excel.util.StyleUtil.buildSXSSFSheetStyle;

/**
 * A context is the main anchorage point of a excel writer.
 *
 * @author jipengfei
 */
public class WriteContext {

    /***
     * The sheet currently written
     */
    private Sheet currentSheet;

    /**
     * current param
     */
    private com.sailvan.excel.metadata.Sheet currentSheetParam;

    /**
     * The sheet currently written's name
     */
    private String currentSheetName;

    /**
     *
     */
    private com.sailvan.excel.metadata.Table currentTable;

    /**
     * Excel type
     */
    private ExcelTypeEnum excelType;

    /**
     * POI Workbook
     */
    private Workbook workbook;

    /**
     * Final output stream
     */
    private OutputStream outputStream;

    /**
     * Written form collection
     */
    private Map<Integer, com.sailvan.excel.metadata.Table> tableMap = new ConcurrentHashMap<Integer, com.sailvan.excel.metadata.Table>();

    /**
     * Cell default head style
     */
    private CellStyle defaultHeadCellStyle;

    /**
     * Cell default content style
     */
    private CellStyle defaultContentCellStyle;

    /**
     * Current table head  style
     */
    private CellStyle currentHeadCellStyle;

    /**
     * Current table content  style
     */
    private CellStyle currentContentCellStyle;

    /**
     * the header attribute of excel
     */
    private ExcelHeadProperty excelHeadProperty;

    private List<String> contentTitle;

    private boolean needHead;

    private WriteHandler afterWriteHandler;

    public WriteHandler getAfterWriteHandler() {
        return afterWriteHandler;
    }

    public WriteContext(InputStream templateInputStream, OutputStream out, ExcelTypeEnum excelType,
                        boolean needHead, WriteHandler afterWriteHandler) throws IOException {
        this.needHead = needHead;
        this.outputStream = out;
        this.afterWriteHandler = afterWriteHandler;
        this.workbook = WorkBookUtil.createWorkBook(templateInputStream, excelType);
        this.defaultHeadCellStyle = StyleUtil.buildDefaultHeadCellStyle(this.workbook);
        this.defaultContentCellStyle = StyleUtil.buildDefaultContentCellStyle(this.workbook);

    }

    /**
     * @param sheet
     */
    public void currentSheet(com.sailvan.excel.metadata.Sheet sheet) {
        if (null == currentSheetParam || currentSheetParam.getSheetNo() != sheet.getSheetNo()) {
            cleanCurrentSheet();
            currentSheetParam = sheet;
            try {
                this.currentSheet = workbook.getSheetAt(sheet.getSheetNo() - 1);
            } catch (Exception e) {
                this.currentSheet = WorkBookUtil.createSheet(workbook, sheet);
                if (null != afterWriteHandler) {
                    this.afterWriteHandler.sheet(sheet.getSheetNo(), currentSheet);
                }
            }
            buildSXSSFSheetStyle((SXSSFSheet) currentSheet, sheet.getColumnWidthMap());
            /** **/
            this.initCurrentSheet(sheet);
        }

    }

    private void initCurrentSheet(com.sailvan.excel.metadata.Sheet sheet) {

        initExcelHeadProperty(getHead(sheet), sheet.getClazz());

        initContentTitle(sheet.getContentTitle());

        initTableStyle(sheet.getTableStyle());

        initTableHead();

    }

    private void cleanCurrentSheet() {
        this.currentSheet = null;
        this.currentSheetParam = null;
        this.excelHeadProperty = null;
        this.contentTitle = null;
        this.currentHeadCellStyle = null;
        this.currentContentCellStyle = null;
        this.currentTable = null;

    }

    /**
     * init excel header
     *
     * @param head
     * @param clazz
     */
    private void initExcelHeadProperty(List<List<String>> head, Class<? extends BaseRowModel> clazz) {
        if (head != null || clazz != null) {
            this.excelHeadProperty = new ExcelHeadProperty(clazz, head);
            this.currentSheetParam.setHead(excelHeadProperty.getHead());
        }
    }

    private void initContentTitle(List<String> contentTitle) {
        if (contentTitle != null) {
            this.contentTitle = contentTitle;
        }
    }

    public void initTableHead() {
        if (needHead && null != excelHeadProperty && !CollectionUtils.isEmpty(excelHeadProperty.getHead())) {
            int startRow = currentSheet.getLastRowNum();
            if (startRow > 0) {
                startRow += 4;
            } else {
                startRow = currentSheetParam.getStartRow();
            }
            addMergedRegionToCurrentSheet(startRow);
            int i = startRow;
            for (; i < this.excelHeadProperty.getRowNum() + startRow; i++) {
                Row row = WorkBookUtil.createRow(currentSheet, i);
                if (null != afterWriteHandler) {
                    this.afterWriteHandler.row(i, row);
                }
                addOneRowOfHeadDataToExcel(row, this.excelHeadProperty.getHeadByRowNum(i - startRow));
            }
        }
    }

    private void addMergedRegionToCurrentSheet(int startRow) {
        for (com.sailvan.excel.metadata.CellRange cellRangeModel : excelHeadProperty.getCellRangeModels()) {
            currentSheet.addMergedRegion(new CellRangeAddress(cellRangeModel.getFirstRow() + startRow,
                cellRangeModel.getLastRow() + startRow,
                cellRangeModel.getFirstCol(), cellRangeModel.getLastCol()));
        }
    }

    private void addOneRowOfHeadDataToExcel(Row row, List<String> headByRowNum) {
        StyleUtil.buildRowStyle(row);
        if (headByRowNum != null && headByRowNum.size() > 0) {
            for (int i = 0; i < headByRowNum.size(); i++) {
                Cell cell = WorkBookUtil.createCell(row, i, getCurrentHeadCellStyle(), headByRowNum.get(i));
                if (null != afterWriteHandler) {
                    this.afterWriteHandler.cell(i, cell, null);
                }
            }
        }
    }

    private void initTableStyle(com.sailvan.excel.metadata.TableStyle tableStyle) {
        if (tableStyle != null) {
            this.currentHeadCellStyle = StyleUtil.buildCellStyle(this.workbook, tableStyle.getTableHeadFont(),
                tableStyle.getTableHeadBackGroundColor());
            this.currentContentCellStyle = StyleUtil.buildCellStyle(this.workbook, tableStyle.getTableContentFont(),
                tableStyle.getTableContentBackGroundColor());
        }
    }

    private void cleanCurrentTable() {
        this.excelHeadProperty = null;
        this.currentHeadCellStyle = null;
        this.currentContentCellStyle = null;
        this.currentTable = null;

    }

    public void currentTable(com.sailvan.excel.metadata.Table table) {
        if (null == currentTable || currentTable.getTableNo() != table.getTableNo()) {
            cleanCurrentTable();
            this.currentTable = table;
            this.initExcelHeadProperty(table.getHead(), table.getClazz());
            this.initTableStyle(table.getTableStyle());
            this.initTableHead();
        }

    }

    public List<List<String>> getHead(com.sailvan.excel.metadata.Sheet sheet) {
        return sheet.getHead() != null ? sheet.getHead() : ListUtil.listWarp(sheet.getSingleHead());
    }

    public ExcelHeadProperty getExcelHeadProperty() {
        return this.excelHeadProperty;
    }

    public List<String> getContentTitle() {
        return contentTitle;
    }

    public boolean needHead() {
        return this.needHead;
    }

    public Sheet getCurrentSheet() {
        return currentSheet;
    }

    public void setCurrentSheet(Sheet currentSheet) {
        this.currentSheet = currentSheet;
    }

    public String getCurrentSheetName() {
        return currentSheetName;
    }

    public void setCurrentSheetName(String currentSheetName) {
        this.currentSheetName = currentSheetName;
    }

    public ExcelTypeEnum getExcelType() {
        return excelType;
    }

    public void setExcelType(ExcelTypeEnum excelType) {
        this.excelType = excelType;
    }

    public OutputStream getOutputStream() {
        return outputStream;
    }

    public CellStyle getCurrentHeadCellStyle() {
        return this.currentHeadCellStyle == null ? defaultHeadCellStyle : this.currentHeadCellStyle;
    }

    public CellStyle getCurrentContentStyle() {
        return this.currentContentCellStyle == null? defaultContentCellStyle : this.currentContentCellStyle;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public com.sailvan.excel.metadata.Sheet getCurrentSheetParam() {
        return currentSheetParam;
    }

    public void setCurrentSheetParam(com.sailvan.excel.metadata.Sheet currentSheetParam) {
        this.currentSheetParam = currentSheetParam;
    }

    public com.sailvan.excel.metadata.Table getCurrentTable() {
        return currentTable;
    }

    public void setCurrentTable(com.sailvan.excel.metadata.Table currentTable) {
        this.currentTable = currentTable;
    }
}


