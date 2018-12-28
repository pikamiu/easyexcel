package com.alibaba.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;

import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.TypeReference;
import com.alibaba.fastjson.parser.Feature;

/**
 * <p>easy excel util </p>
 *
 * @author wujiaming
 */
public class EasyExcelUtil {

    public static List<List<String>> read(String path) {
        List<List<String>> lists = new ArrayList<List<String>>();
        try {
            lists = read(getFileInputStream(path), 1, 0);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return lists;
    }

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

    public static <T> List<T> read(String path, Class<? extends BaseRowModel> tClass) throws FileNotFoundException {
        return read(getFileInputStream(path), 1, 1, tClass);
    }

    public static <T> List<T> read(String path, int sheetNo, int headLineMun, Class<? extends BaseRowModel> tClass) throws FileNotFoundException {
        return read(getFileInputStream(path), sheetNo, headLineMun, tClass);
    }

    @SuppressWarnings("unchecked")
    public static <T> List<T> read(InputStream in, int sheetNo, int headLineMun, Class<? extends BaseRowModel> tClass) {
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

    public static void writeWithMultiHead(List<List<Object>> lists, List<List<String>> head, String path) {
        OutputStream out = null;
        try {
            out = new FileOutputStream(path);
            ExcelWriter writer = EasyExcelFactory.getWriter(out);
            writer.writeObject(lists, new Sheet(1, 1, head));
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
            List<List<String>> head = listTransform(jsonParse(lists.get(0)));
            if (sheet.getHead() == null) {
                sheet.setHead(head);
            }
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

    public static ExcelWriter writeStream(String path) throws FileNotFoundException {
        OutputStream out = new FileOutputStream(path);
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
        LinkedHashMap<String, Object> rootStr = JSON.parseObject(str, new TypeReference<LinkedHashMap<String, Object>>() {}, Feature.OrderedField);
        return new ArrayList<String>(rootStr.keySet());
    }
}
