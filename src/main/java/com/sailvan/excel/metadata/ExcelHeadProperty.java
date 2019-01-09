package com.sailvan.excel.metadata;

import com.sailvan.excel.annotation.ExcelColumnNum;
import com.sailvan.excel.annotation.ExcelProperty;
import com.sailvan.excel.metadata.typeconvertor.TypeConvertor;

import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.*;

/**
 * Define the header attribute of excel
 *
 * @author jipengfei
 */
public class ExcelHeadProperty {

    /**
     * Custom class
     */
    private Class<? extends BaseRowModel> headClazz;

    /**
     * A two-dimensional array describing the header
     */
    private List<List<String>> head;

    /**
     * Attributes described by the header
     */
    private List<ExcelColumnProperty> columnPropertyList = new ArrayList<ExcelColumnProperty>();

    /**
     * Attributes described by the header
     */
    private Map<Integer, ExcelColumnProperty> excelColumnPropertyMap1 = new HashMap<Integer, ExcelColumnProperty>();

    public ExcelHeadProperty(Class<? extends BaseRowModel> headClazz, List<List<String>> head) {
        this.headClazz = headClazz;
        this.head = head;
        initColumnProperties();
    }

    private void initColumnProperties() {
        if (this.headClazz != null) {
            boolean globalAnnotation = initClassTypeAnnotation();
            Map<String, Integer> orderMap = initOrderType(globalAnnotation);

            List<Field> fieldList = new ArrayList<Field>();
            Class tempClass = this.headClazz;
            /*
            When the parent class is null, it indicates that the parent class (Object class) has reached
            the top level
             */
            while (tempClass != null) {
                fieldList.addAll(Arrays.asList(tempClass.getDeclaredFields()));
                // Get the parent class and give it to yourself
                if (!globalAnnotation)
                    tempClass = tempClass.getSuperclass();
                else
                    tempClass = null;
            }
            removeIgnore(fieldList);
            for (int i = 0; i < fieldList.size(); i++) {
                if (globalAnnotation)
                    initOneColumnPropertyWithType(fieldList.get(i), i, orderMap);
                else
                    initOneColumnProperty(fieldList.get(i));
            }
            // sort the columns
            Collections.sort(columnPropertyList);

            List<List<String>> headList = new ArrayList<List<String>>();
            if (head == null || head.size() == 0) {
                for (ExcelColumnProperty excelColumnProperty : columnPropertyList) {
                    headList.add(excelColumnProperty.getHead());
                }
                this.head = headList;
            }
        }
    }

    private boolean initClassTypeAnnotation() {
        return headClazz.getAnnotation(ExcelProperty.class) != null;
    }

    private Map<String, Integer> initOrderType(boolean globalAnnotation) {
        Map<String, Integer> map = new LinkedHashMap<String, Integer>();
        if (globalAnnotation) {
            ExcelProperty excelProperty = headClazz.getAnnotation(ExcelProperty.class);
            String[] orders = excelProperty.orders();
            for (int i = 0; i < orders.length; i++) {
                map.put(orders[i], i);
            }
        }

        return map;
    }

    private void removeIgnore(List<Field> fields) {
        Iterator<Field> it = fields.iterator();
        while (it.hasNext()) {
            ExcelProperty p = it.next().getAnnotation(ExcelProperty.class);
            if (p != null && p.ignore())
                it.remove();
        }
    }

    private void initOneColumnPropertyWithType(Field f, Integer indexSelf, Map<String, Integer> orderMap) {
        ExcelProperty p = f.getAnnotation(ExcelProperty.class);
        ExcelColumnProperty excelHeadProperty;
        indexSelf = orderMap.isEmpty() ? indexSelf : orderMap.get(f.getName());
        if (indexSelf == null) {
            throw new RuntimeException("The sequence of order Annotation incomplete");
        }
        if (p != null) {
            if (p.ignore()) return;
            excelHeadProperty = new ExcelColumnProperty();
            List<String> head = p.value().length != 0 ? Arrays.asList(p.value()) : Collections.singletonList(f.getName());
            excelHeadProperty.setField(f);
            excelHeadProperty.setHead(head);
            excelHeadProperty.setIndex(indexSelf);
            excelHeadProperty.setFormat(p.format());
            addConvertor(p, excelHeadProperty);
            excelColumnPropertyMap1.put(indexSelf, excelHeadProperty);
        } else {
            excelHeadProperty = new ExcelColumnProperty();
            excelHeadProperty.setField(f);
            excelHeadProperty.setHead(Collections.singletonList(f.getName()));
            excelHeadProperty.setIndex(indexSelf);
            excelHeadProperty.setFormat("");
            excelColumnPropertyMap1.put(indexSelf, excelHeadProperty);
        }
        this.columnPropertyList.add(excelHeadProperty);
    }

    /**
     * @param f
     */
    private void initOneColumnProperty(Field f) {
        ExcelProperty p = f.getAnnotation(ExcelProperty.class);
        ExcelColumnProperty excelHeadProperty = null;
        if (p != null) {
            excelHeadProperty = new ExcelColumnProperty();
            excelHeadProperty.setField(f);
            excelHeadProperty.setHead(Arrays.asList(p.value()));
            excelHeadProperty.setIndex(p.index());
            excelHeadProperty.setFormat(p.format());
            addConvertor(p, excelHeadProperty);
            excelColumnPropertyMap1.put(p.index(), excelHeadProperty);
        } else {
            ExcelColumnNum columnNum = f.getAnnotation(ExcelColumnNum.class);
            if (columnNum != null) {
                excelHeadProperty = new ExcelColumnProperty();
                excelHeadProperty.setField(f);
                excelHeadProperty.setIndex(columnNum.value());
                excelHeadProperty.setFormat(columnNum.format());
                excelColumnPropertyMap1.put(columnNum.value(), excelHeadProperty);
            }
        }
        if (excelHeadProperty != null) {
            this.columnPropertyList.add(excelHeadProperty);
        }

    }

    private void addConvertor(ExcelProperty p, ExcelColumnProperty excelHeadProperty){
        try {
            if (!p.convertor().equals(TypeConvertor.class) && (p.convertor().isInterface() || Modifier.isAbstract(p.convertor().getModifiers())))
                throw new RuntimeException("convertor type is wrong");
            else if (!p.convertor().equals(TypeConvertor.class))
                excelHeadProperty.setConverter((TypeConvertor) p.convertor().newInstance());
        } catch (InstantiationException e) {
            throw new RuntimeException("InstantiationException is happen : {}" , e);
        } catch (IllegalAccessException e) {
            throw new RuntimeException("IllegalAccessException is happen : {}" , e);
        }
    }

    /**
     *
     */
    public void appendOneRow(List<String> row) {

        for (int i = 0; i < row.size(); i++) {
            List<String> list;
            if (head.size() <= i) {
                list = new ArrayList<String>();
                head.add(list);
            } else {
                list = head.get(0);
            }
            list.add(row.get(i));
        }

    }

    /**
     * @param columnNum
     * @return
     */
    public ExcelColumnProperty getExcelColumnProperty(int columnNum) {
        return excelColumnPropertyMap1.get(columnNum);
    }

    public Class getHeadClazz() {
        return headClazz;
    }

    public void setHeadClazz(Class headClazz) {
        this.headClazz = headClazz;
    }

    public List<List<String>> getHead() {
        return this.head;
    }

    public void setHead(List<List<String>> head) {
        this.head = head;
    }

    public List<ExcelColumnProperty> getColumnPropertyList() {
        return columnPropertyList;
    }

    public void setColumnPropertyList(List<ExcelColumnProperty> columnPropertyList) {
        this.columnPropertyList = columnPropertyList;
    }

    /**
     * Calculate all cells that need to be merged
     *
     * @return cells that need to be merged
     */
    public List<CellRange> getCellRangeModels() {
        List<CellRange> cellRanges = new ArrayList<CellRange>();
        for (int i = 0; i < head.size(); i++) {
            List<String> columnValues = head.get(i);
            for (int j = 0; j < columnValues.size(); j++) {
                int lastRow = getLastRangNum(j, columnValues.get(j), columnValues);
                int lastColumn = getLastRangNum(i, columnValues.get(j), getHeadByRowNum(j));
                if ((lastRow > j || lastColumn > i) && lastRow >= 0 && lastColumn >= 0) {
                    cellRanges.add(new CellRange(j, lastRow, i, lastColumn));
                }
            }
        }
        return cellRanges;
    }

    public List<String> getHeadByRowNum(int rowNum) {
        List<String> l = new ArrayList<String>(head.size());
        for (List<String> list : head) {
            if (list.size() > rowNum) {
                l.add(list.get(rowNum));
            } else {
                l.add(list.get(list.size() - 1));
            }
        }
        return l;
    }

    /**
     * Get the last consecutive string position
     *
     * @param j      current value position
     * @param value  value content
     * @param values values
     * @return the last consecutive string position
     */
    private int getLastRangNum(int j, String value, List<String> values) {
        if (value == null) {
            return -1;
        }
        if (j > 0) {
            String preValue = values.get(j - 1);
            if (value.equals(preValue)) {
                return -1;
            }
        }
        int last = j;
        for (int i = last + 1; i < values.size(); i++) {
            String current = values.get(i);
            if (value.equals(current)) {
                last = i;
            } else {
                // if i>j && !value.equals(current) Indicates that the continuous range is exceeded
                if (i > j) {
                    break;
                }
            }
        }
        return last;

    }

    public int getRowNum() {
        int headRowNum = 0;
        for (List<String> list : head) {
            if (list != null && list.size() > 0) {
                if (list.size() > headRowNum) {
                    headRowNum = list.size();
                }
            }
        }
        return headRowNum;
    }

}
