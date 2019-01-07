package com.sailvan.excel;

import com.sailvan.excel.event.WriteHandler;
import com.sailvan.excel.metadata.BaseRowModel;
import com.sailvan.excel.metadata.Sheet;
import com.sailvan.excel.metadata.Table;
import com.sailvan.excel.parameter.GenerateParam;
import com.sailvan.excel.support.ExcelTypeEnum;
import com.sailvan.excel.write.ExcelBuilder;
import com.sailvan.excel.write.ExcelBuilderImpl;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

/**
 * Excel Writer This tool is used to write data out to Excel via POI.
 * This object can perform the following two functions.
 * <pre>
 *    1. Create a new empty Excel workbook, write the data to the stream after the data is filled.
 *    2. Edit existing Excel, write the original Excel file, or write it to other places.}
 * </pre>
 * @author jipengfei
 */
public class ExcelWriter {

    private ExcelBuilder excelBuilder;

    /**
     * Create new writer
     * @param outputStream the java OutputStream you wish to write the data to
     * @param typeEnum 03 or 07
     */
    public ExcelWriter(OutputStream outputStream, ExcelTypeEnum typeEnum) {
        this(outputStream, typeEnum, true);
    }

    @Deprecated
    private Class<? extends BaseRowModel> objectClass;

    /**
     * @param generateParam
     */
    @Deprecated
    public ExcelWriter(GenerateParam generateParam) {
        this(generateParam.getOutputStream(), generateParam.getType(), true);
        this.objectClass = generateParam.getClazz();
    }

    /**
     *
     * Create new writer
     * @param outputStream the java OutputStream you wish to write the data to
     * @param typeEnum 03 or 07
     * @param needHead Do you need to write the header to the file?
     */
    public ExcelWriter(OutputStream outputStream, ExcelTypeEnum typeEnum, boolean needHead) {
        excelBuilder = new ExcelBuilderImpl(null, outputStream, typeEnum, needHead, null);
    }

    /**
     *  Create new writer
     * @param templateInputStream Append data after a POI file ,Can be null（the template POI filesystem that contains the Workbook stream)
     * @param outputStream the java OutputStream you wish to write the data to
     * @param typeEnum 03 or 07
     */
    public ExcelWriter(InputStream templateInputStream, OutputStream outputStream, ExcelTypeEnum typeEnum,Boolean needHead) {
        excelBuilder = new ExcelBuilderImpl(templateInputStream,outputStream, typeEnum, needHead, null);
    }


    /**
     *  Create new writer
     * @param templateInputStream Append data after a POI file ,Can be null（the template POI filesystem that contains the Workbook stream)
     * @param outputStream the java OutputStream you wish to write the data to
     * @param typeEnum 03 or 07
     * @param writeHandler User-defined callback
     */
    public ExcelWriter(InputStream templateInputStream, OutputStream outputStream, ExcelTypeEnum typeEnum, Boolean needHead,
                       WriteHandler writeHandler) {
        excelBuilder = new ExcelBuilderImpl(templateInputStream,outputStream, typeEnum, needHead,writeHandler);
    }

    /**
     * Write data to a sheet
     * @param data Data to be written
     * @param sheet Write to this sheet
     * @return this current writer
     */
    public ExcelWriter writeBean(List<? extends BaseRowModel> data, Sheet sheet) {
        excelBuilder.addContent(data, sheet);
        return this;
    }


    /**
     * Write data to a sheet
     * @param data Data to be written
     * @return this current writer
     */
    @Deprecated
    public ExcelWriter write(List data) {
        if (objectClass != null) {
            return this.writeBean(data,new Sheet(1,0,objectClass));
        }else {
            return this.writeString(data,new Sheet(1,0,objectClass));

        }
    }

    /**
     *
     * Write data to a sheet
     * @param data Data to be written
     * @param sheet Write to this sheet
     * @return this
     */
    public ExcelWriter writeObject(List<List<Object>> data, Sheet sheet) {
        excelBuilder.addContent(data, sheet);
        return this;
    }


    /**
     * Write data to a sheet, main use to write json data
     * @param data Data to be written
     * @param sheet Write to this sheet
     * @return this
     */
    public ExcelWriter writeJson(List<String> data, Sheet sheet) {
        if (data.size() <= 0)
            throw new IllegalArgumentException("write json size can't empty");
        excelBuilder.addHead(data.get(0), sheet);
        excelBuilder.addContent(data, sheet);
        return this;
    }


    /**
     * Write data to a sheet
     * @param data  Data to be written
     * @param sheet Write to this sheet
     * @return this
     */
    public ExcelWriter writeString(List<List<String>> data, Sheet sheet) {
        excelBuilder.addContent(data, sheet);
        return this;
    }

    /**
     * Write data to a sheet
     * @param data  Data to be written
     * @param sheet Write to this sheet
     * @param table Write to this table
     * @return this
     */
    public ExcelWriter writeBean(List<? extends BaseRowModel> data, Sheet sheet, Table table) {
        excelBuilder.addContent(data, sheet, table);
        return this;
    }

    /**
     * Write data to a sheet
     * @param data  Data to be written
     * @param sheet Write to this sheet
     * @param table Write to this table
     * @return this
     */
    public ExcelWriter writeString(List<List<String>> data, Sheet sheet, Table table) {
        excelBuilder.addContent(data, sheet, table);
        return this;
    }

    public ExcelWriter writeMap(List<Map<String, Object>> data, Sheet sheet) {
        excelBuilder.addContent(data, sheet);
        return this;
    }

    /**
     * Merge Cells，Indexes are zero-based.
     *
     * @param firstRow Index of first row
     * @param lastRow Index of last row (inclusive), must be equal to or larger than {@code firstRow}
     * @param firstCol Index of first column
     * @param lastCol Index of last column (inclusive), must be equal to or larger than {@code firstCol}
     */
    public ExcelWriter merge(int firstRow, int lastRow, int firstCol, int lastCol){
        excelBuilder.merge(firstRow,lastRow,firstCol,lastCol);
        return this;
    }

    /**
     * Write data to a sheet
     * @param data  Data to be written
     * @param sheet Write to this sheet
     * @param table Write to this table
     * @return
     */
    public ExcelWriter writeObject(List<List<Object>> data, Sheet sheet, Table table) {
        excelBuilder.addContent(data, sheet, table);
        return this;
    }

    /**
     * Close IO
     */
    public void finish() {
        excelBuilder.finish();
    }
}
