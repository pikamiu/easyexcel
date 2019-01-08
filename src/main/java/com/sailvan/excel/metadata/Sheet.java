package com.sailvan.excel.metadata;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 *
 * @author jipengfei
 */
public class Sheet {

    /**
     * the head line number
     */
    private int headLineMun;

    /**
     * Starting from 1
     */
    private int sheetNo;

    private String sheetName;

    private Class<? extends BaseRowModel> clazz;

    /**
     * multi line head
     */
    private List<List<String>> head;

    /**
     * one line head
     */
    private List<String> singleHead;

    /**
     * content title
     */
    private List<String> contentTitle;

    private TableStyle tableStyle;

    /**
     * column with
     */
    private Map<Integer,Integer> columnWidthMap = new HashMap<Integer, Integer>();

    /**
     * boolean autoWidth, it's notice that it will make program more slowly
     */
    private Boolean autoWidth = Boolean.FALSE;

    private int startRow = 0;


    public Sheet(int sheetNo) {
        this.sheetNo = sheetNo;
    }

    public Sheet(int sheetNo, int headLineMun) {
        this.sheetNo = sheetNo;
        this.headLineMun = headLineMun;
    }

    public Sheet (int sheetNo, int headLineMun, List<List<String>> head) {
        this.sheetNo = sheetNo;
        this.headLineMun = headLineMun;
        this.head = head;
    }

    public Sheet(int sheetNo, int headLineMun, Class<? extends BaseRowModel> clazz) {
        this.sheetNo = sheetNo;
        this.headLineMun = headLineMun;
        this.clazz = clazz;
    }

    public Sheet(int sheetNo, int headLineMun, Class<? extends BaseRowModel> clazz, String sheetName,
                 List<List<String>> head) {
        this.sheetNo = sheetNo;
        this.clazz = clazz;
        this.headLineMun = headLineMun;
        this.sheetName = sheetName;
        this.head = head;
    }

    public List<List<String>> getHead() {
        return head;
    }

    public void setHead(List<List<String>> head) {
        this.head = head;
    }

    public List<String> getSingleHead() {
        return singleHead;
    }

    public void setSingleHead(List<String> singleHead) {
        this.singleHead = singleHead;
    }

    public List<String> getContentTitle() {
        return contentTitle;
    }

    public void setContentTitle(List<String> contentTitle) {
        this.contentTitle = contentTitle;
    }

    public Class<? extends BaseRowModel> getClazz() {
        return clazz;
    }

    public void setClazz(Class<? extends BaseRowModel> clazz) {
        this.clazz = clazz;
        if (headLineMun == 0) {
            this.headLineMun = 1;
        }
    }

    public int getHeadLineMun() {
        return headLineMun;
    }

    public void setHeadLineMun(int headLineMun) {
        this.headLineMun = headLineMun;
    }

    public int getSheetNo() {
        return sheetNo;
    }

    public void setSheetNo(int sheetNo) {
        this.sheetNo = sheetNo;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public TableStyle getTableStyle() {
        return tableStyle;
    }

    public void setTableStyle(TableStyle tableStyle) {
        this.tableStyle = tableStyle;
    }



    public Map<Integer, Integer> getColumnWidthMap() {
        return columnWidthMap;
    }

    public void setColumnWidthMap(Map<Integer, Integer> columnWidthMap) {
        this.columnWidthMap = columnWidthMap;
    }

    @Override
    public String toString() {
        return "Sheet{" +
            "headLineMun=" + headLineMun +
            ", sheetNo=" + sheetNo +
            ", sheetName='" + sheetName + '\'' +
            ", clazz=" + clazz +
            ", head=" + head +
                ", singleHead=" + singleHead +
            ", tableStyle=" + tableStyle +
            ", columnWidthMap=" + columnWidthMap +
            '}';
    }

    public Boolean getAutoWidth() {
        return autoWidth;
    }

    public void setAutoWidth(Boolean autoWidth) {
        this.autoWidth = autoWidth;
    }

    public int getStartRow() {
        return startRow;
    }

    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }



    private Sheet(Builder builder) {
        this.headLineMun = builder.headLineMun;
        this.sheetNo = builder.sheetNo;
        this.sheetName = builder.sheetName;
        this.clazz = builder.clazz;
        this.head = builder.head;
        this.singleHead = builder.singleHead;
        this.contentTitle = builder.contentTitle;
        this.tableStyle = builder.tableStyle;
        this.columnWidthMap = builder.columnWidthMap;
        this.autoWidth = builder.autoWidth;
        this.startRow = builder.startRow;
    }

    public static class Builder implements Builders<Sheet> {
        private int headLineMun = 0;
        private int sheetNo = 1;
        private String sheetName;
        private Class<? extends BaseRowModel> clazz;
        private List<List<String>> head;
        private List<String> singleHead;
        private List<String> contentTitle;
        private TableStyle tableStyle;
        private Map<Integer,Integer> columnWidthMap = new HashMap<Integer, Integer>();
        private Boolean autoWidth = Boolean.FALSE;
        private int startRow = 0;

        public Builder sheetNo(int val) {
            sheetNo = val;
            return this;
        }
        public Builder headLineMun(int val) {
            headLineMun = val;
            return this;
        }
        public Builder sheetName(String val) {
            sheetName = val;
            return this;
        }
        public Builder clazz(Class<? extends BaseRowModel> val) {
            clazz = val;
            return this;
        }
        public Builder head(List<List<String>> val) {
            head = val;
            return this;
        }
        public Builder singleHead(List<String> val) {
            singleHead = val;
            return this;
        }
        public Builder contentTitle(List<String> val) {
            contentTitle = val;
            return this;
        }
        public Builder tableStyle(TableStyle val) {
            tableStyle = val;
            return this;
        }
        public Builder columnWidthMap(Map<Integer,Integer> val) {
            columnWidthMap = val;
            return this;
        }
        public Builder autoWidth(boolean val) {
            autoWidth = val;
            return this;
        }
        public Builder startRow(int val) {
            startRow = val;
            return this;
        }

        @Override
        public Sheet build() {
            return new Sheet(this);
        }
    }

}
