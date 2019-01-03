package com.alibaba.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;

import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Sheet.SheetBuilder;
import com.alibaba.excel.metadata.WriteInfo;
import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.TypeReference;
import com.alibaba.fastjson.parser.Feature;

/**
 * <p>easy excel util </p>
 *
 * @author wujiaming
 */
public class EasyExcelUtil {

    /**
     * According to the path read excel
     * default sheetNumber 1 and headLineMun 0
     *
     * @param path the file path
     * @return data result
     */
    public static List<List<String>> read(String path) {
        List<List<String>> lists = new ArrayList<List<String>>();
        try {
            lists = read(getFileInputStream(path), 1, 1);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return lists;
    }

    /**
     * read excel
     *
     * @param path        file path
     * @param sheetNo     sheetNumber
     * @param headLineMun headLineNumber
     * @return data result
     */
    public static List<List<String>> read(String path, int sheetNo, int headLineMun) {
        List<List<String>> lists = new ArrayList<List<String>>();
        try {
            lists = read(getFileInputStream(path), sheetNo, headLineMun);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return lists;
    }

    /**
     * read excel
     *
     * @param in          stream
     * @param sheetNo     sheetNumber
     * @param headLineMun headLineNumber it's use to judge context start line
     * @return data result
     */
    public static List<List<String>> read(InputStream in, int sheetNo, int headLineMun) {
        List<List<String>> lists = new ArrayList<List<String>>();
        try {
            lists = EasyExcelFactory.read2ListString(in, new Sheet(sheetNo, headLineMun));
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (in != null) {
                try {
                    in.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return lists;
    }

    /**
     * read excel by java bean, the bean must extend BaseRowModel, This is a tag interface.
     * you can use ExcelProperty annotation design the java bean
     * default sheetNumber 1, headLineMun 1,
     *
     * @param path   file path
     * @param tClass java bean what extend the BaseRowModel
     * @param <T>    T
     * @return list<bean>
     * @see com.alibaba.excel.annotation.ExcelProperty
     */
    public static <T> List<T> read(String path, Class<? extends BaseRowModel> tClass) {
        try {
            return read(getFileInputStream(path), tClass, 1, 1);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * read excel by java bean, the bean must extend BaseRowModel, This is a tag interface.
     * you can use ExcelProperty annotation design the java bean
     * <p>
     * Please note that the headLineMun must greater than the head line count,because the head dataType maybe different
     * to context data type, It may throw ClassCastException.
     *
     * @param path        file path
     * @param tClass      java bean what extend the BaseRowModel
     * @param sheetNo     sheetNumber
     * @param headLineMun headLineNumber it's use to judge context start line
     * @param <T>         T
     * @return list<bean>
     * @see com.alibaba.excel.annotation.ExcelProperty
     */
    public static <T> List<T> read(String path, Class<? extends BaseRowModel> tClass, int sheetNo, int headLineMun) {
        try {
            return read(getFileInputStream(path), tClass, sheetNo, headLineMun);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * read excel by java bean with InputStream
     *
     * @param in          InputStream
     * @param tClass      java bean what extend the BaseRowModel
     * @param sheetNo     sheetNumber
     * @param headLineMun headLineNumber
     * @param <T>         T
     * @return list<bean>
     */
    @SuppressWarnings("unchecked")
    public static <T> List<T> read(InputStream in, Class<? extends BaseRowModel> tClass, int sheetNo, int headLineMun) {
        List<T> lists = new ArrayList<T>();
        try {
            lists = (List<T>) EasyExcelFactory.read(in, new Sheet(sheetNo, headLineMun, tClass));
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (in != null) {
                try {
                    in.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return lists;
    }

    /**
     * custom read excel, it's need to Override AnalysisEventListener
     *
     * @param path     file path
     * @param listener AnalysisEventListener what need to Override
     * @throws FileNotFoundException if file not exist
     * @see com.alibaba.excel.event.AnalysisEventListener
     */
    public static void read(String path, AnalysisEventListener listener) {
        try {
            EasyExcelFactory.readBySax(getFileInputStream(path), new Sheet(1, 1), listener);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }

    /**
     * custom read excel, it's need to Override AnalysisEventListener
     *
     * @param path        file path
     * @param sheetNo     SheetNumber
     * @param headLineMun headLine
     * @param listener    AnalysisEventListener what need to Override
     * @throws FileNotFoundException if file not exist
     * @see com.alibaba.excel.event.AnalysisEventListener
     */
    public static void read(String path, int sheetNo, int headLineMun, AnalysisEventListener listener) {
        try {
            EasyExcelFactory.readBySax(getFileInputStream(path), new Sheet(sheetNo, headLineMun), listener);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }

    /**
     * custom read excel, it's need to Override AnalysisEventListener
     *
     * @param path     file path
     * @param sheet    Sheet
     * @param listener AnalysisEventListener what need to Override
     * @see com.alibaba.excel.event.AnalysisEventListener
     */
    public static void read(InputStream in, Sheet sheet, AnalysisEventListener listener) {
        EasyExcelFactory.readBySax(in, sheet, listener);
    }


    /**
     * write excel
     *
     * @param lists      data list
     * @param singleHead one line head
     * @param path       file out path
     */
    public static void write(List<List<Object>> lists, List<String> singleHead, String path) {
        OutputStream out = null;
        try {
            out = new FileOutputStream(path);
            ExcelWriter writer = EasyExcelFactory.getWriter(out);
            writer.writeObject(lists, new Sheet(1, 1, listTransform(singleHead)));
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * write excel
     *
     * @param lists     data list
     * @param multiHead combine head title
     * @param path      file path
     */
    public static void writeWithMultiHead(List<List<Object>> lists, List<List<String>> multiHead, String path) {
        OutputStream out = null;
        try {
            out = new FileOutputStream(path);
            ExcelWriter writer = EasyExcelFactory.getWriter(out);
            writer.writeObject(lists, new Sheet(1, 1, multiHead));
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static void write(List<List<Object>> lists, Sheet sheet, String path) {
        OutputStream out = null;
        try {
            out = new FileOutputStream(path);
            ExcelWriter writer = EasyExcelFactory.getWriter(out);
            writer.writeObject(lists, sheet);
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static void writeByBean(List<? extends BaseRowModel> models, Class<? extends BaseRowModel> tClass, String path) {
        OutputStream out = null;
        try {
            out = new FileOutputStream(path);
            ExcelWriter writer = EasyExcelFactory.getWriter(out);
            writer.writeBean(models, new Sheet(0, 1, tClass));
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static void writeByBean(List<? extends BaseRowModel> models, Sheet sheet, String path) {
        OutputStream out = null;
        try {
            out = new FileOutputStream(path);
            ExcelWriter writer = EasyExcelFactory.getWriter(out);
            writer.writeBean(models, sheet);
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static void writeByJson(List<String> lists, String path) {
        if (lists.size() <= 0) {
            throw new IllegalArgumentException("data can't be empty");
        }
        OutputStream out = null;
        try {
            out = new FileOutputStream(path);
            ExcelWriter writer = EasyExcelFactory.getWriter(out);
            List<List<String>> head = listTransform(jsonParse(lists.get(0)));
            writer.writeJson(lists, new Sheet(1, 1, head));
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static void writeByJson(List<String> lists, Sheet sheet, String path) {
        if (lists.size() <= 0) {
            throw new IllegalArgumentException("data can't be empty");
        }
        OutputStream out = null;
        try {
            out = new FileOutputStream(path);
            ExcelWriter writer = EasyExcelFactory.getWriter(out);
            writer.writeJson(lists, sheet);
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static void write(WriteInfo writeInfo, String path) {
        OutputStream out = null;
        try {
            out = new FileOutputStream(path);
            ExcelWriter writer = EasyExcelFactory.getWriter(out);
            writer.writeMap(writeInfo.getContentList(),
                    new SheetBuilder()
                            .sheetNo(1)
                            .headLineMun(1)
                            .sheetName(writeInfo.getSheetName())
                            .singleHead(Arrays.asList(writeInfo.getTitle()))
                            .contentTitle(Arrays.asList(writeInfo.getContentTitle()))
                            .build());
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static void write(List<WriteInfo> list, String path) {
        OutputStream out = null;
        try {
            out = new FileOutputStream(path);
            ExcelWriter writer = EasyExcelFactory.getWriter(out);
            for (int i = 0; i < list.size(); i++) {
                WriteInfo writeInfo = list.get(i);
                writer.writeMap(writeInfo.getContentList(),
                        new SheetBuilder()
                                .sheetNo(i + 1)
                                .headLineMun(1)
                                .sheetName(writeInfo.getSheetName())
                                .singleHead(Arrays.asList(writeInfo.getTitle()))
                                .contentTitle(Arrays.asList(writeInfo.getContentTitle()))
                                .build());
            }
            writer.finish();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static ExcelWriter writeStream(String path) {
        OutputStream out = null;
        try {
            out = new FileOutputStream(path);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return EasyExcelFactory.getWriter(out);
    }


    public static InputStream getResourcesFileInputStream(String fileName) {
        return Thread.currentThread().getContextClassLoader().getResourceAsStream("" + fileName);
    }

    public static InputStream getFileInputStream(String fileName) throws FileNotFoundException {
        return new FileInputStream(fileName);
    }

    private static List<List<String>> listTransform(List<String> list) {
        List<List<String>> l = new ArrayList<List<String>>();
        for (String s : list) {
            l.add(Collections.singletonList(s));
        }
        return l;
    }

    private static List<String> jsonParse(String str) {
        LinkedHashMap<String, Object> rootStr = JSON.parseObject(str, new TypeReference<LinkedHashMap<String, Object>>() {
        }, Feature.OrderedField);
        return new ArrayList<String>(rootStr.keySet());
    }
}
