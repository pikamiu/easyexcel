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
    T serialize(String s);

    String deserialize(T s);
}
