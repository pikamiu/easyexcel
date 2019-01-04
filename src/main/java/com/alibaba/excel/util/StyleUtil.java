package com.alibaba.excel.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.util.Map;

/**
 * @author jipengfei
 */
public class StyleUtil {

    public static void buildRowSttle(Row row) {
        row.setHeight((short) 600);
    }

    /**
     *
     * @param workbook
     * @return
     */
    public static CellStyle buildDefaultHeadCellStyle(Workbook workbook) {
        XSSFCellStyle newCellStyle = (XSSFCellStyle) workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName("宋体");
        font.setFontHeight((short) 220);
        newCellStyle.setFont(font);

        newCellStyle.setLocked(true);
        newCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        newCellStyle.setAlignment(HorizontalAlignment.CENTER);
        newCellStyle.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        newCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFColor borderColor = new XSSFColor(new java.awt.Color(148, 138, 84));
        newCellStyle.setBorderLeft(BorderStyle.THIN);
        newCellStyle.setBorderRight(BorderStyle.THIN);
        newCellStyle.setLeftBorderColor(borderColor);
        newCellStyle.setRightBorderColor(borderColor);

        return newCellStyle;
    }

    public static CellStyle buildDefaultContentCellStyle(Workbook workbook) {
        CellStyle newCellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName("宋体");
        font.setFontHeightInPoints((short)11);
        font.setBold(false);
        newCellStyle.setFont(font);

        newCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        newCellStyle.setAlignment(HorizontalAlignment.CENTER);

        return newCellStyle;
    }

    /**
     *
     * @param workbook
     * @param f
     * @param indexedColors
     * @return
     */
    public static CellStyle buildCellStyle(Workbook workbook, com.alibaba.excel.metadata.Font f,
                                           IndexedColors indexedColors) {
        CellStyle cellStyle = buildDefaultHeadCellStyle(workbook);
        if (f != null) {
            Font font = workbook.createFont();
            font.setFontName(f.getFontName());
            font.setFontHeightInPoints(f.getFontHeightInPoints());
            font.setBold(f.isBold());
            cellStyle.setFont(font);
        }
        if (indexedColors != null) {
            cellStyle.setFillForegroundColor(indexedColors.getIndex());
        }
        return cellStyle;
    }

    public static Sheet buildSheetStyle(Sheet currentSheet, Map<Integer, Integer> sheetWidthMap){
        currentSheet.setDefaultColumnWidth(20);
        for (Map.Entry<Integer, Integer> entry : sheetWidthMap.entrySet()) {
            currentSheet.setColumnWidth(entry.getKey(), entry.getValue());
        }
        return currentSheet;
    }


}
