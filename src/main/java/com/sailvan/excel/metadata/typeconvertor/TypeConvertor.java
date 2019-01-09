package com.sailvan.excel.metadata.typeconvertor;

/**
 * <p>type convert</p>
 *
 * Custom object and cell data conversion
 * You can modify a cell string, and you can modify data in a converter
 *
 * @author wujiaming
 */
public interface TypeConvertor<T> {

    /**
     * Custom convert data to String
     */
    String convert(T s);

    /**
     * reversal data from the excel
     */
    T reversal(String s);

    /**
     * it's main use to clear data
     */
    void doAfter();
}
