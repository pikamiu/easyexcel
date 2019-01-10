package com.sailvan.easyexcel.test;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import com.alibaba.fastjson.JSONObject;
import com.sailvan.excel.EasyExcelUtil;
import com.sailvan.excel.annotation.ExcelProperty;
import com.sailvan.excel.context.AnalysisContext;
import com.sailvan.excel.event.AnalysisEventListener;
import com.sailvan.excel.event.WriteHandler;
import com.sailvan.excel.metadata.BaseRowModel;
import com.sailvan.excel.metadata.WriteInfo;
import com.sailvan.excel.metadata.typeconvertor.AccountTypeConvertor;
import com.sailvan.excel.metadata.typeconvertor.JsonTypeConvertor;

/**
 * <p>easy excel demo</p>
 *
 * easyExcel 核心功能
 *
 * 读任意大小的03、07版Excel不会OOM
 * 读Excel自动通过注解，把结果映射为java模型
 * 读Excel时候是否对Excel内容做trim()增加容错
 * 读Excel支持自定义行级回调，每读一行数据后进行自定义数据操作
 *
 * 写任意大07版Excel不会OOM
 * 写Excel通过注解将JavaBean自动写入Excel
 * 写Excel可以自定义Excel样式 如：字体，加粗，表头颜色，数据内容颜色
 * 写Excel支持sheet，row，cell级别的写入回调，可高度自定义写入样式
 * Sheet提供多个自定义接口
 *
 * 写入效率较之前提升30%左右
 *
 * @author wujiaming
 * @version 1.0
 * @since 2019/1/6 19:50
 */
public class DemoTest {

    /**
     * 常用读取excel方式
     */
    @Test
    public void read() {
        List<List<String>> list = EasyExcelUtil.read("C:\\Users\\Draher\\Desktop\\sku_status_import_template.xlsx");
        System.out.println(list.size());
    }

    /**
     * 通过model读取excel
     */
    @Test
    public void readByModel() {
        List<SkuStatusTO> list = EasyExcelUtil.read("C:\\Users\\Draher\\Desktop\\writeModel.xlsx", SkuStatusTO.class);
        System.out.println(list.size());
    }

    /**
     * 自定义回调样式，实现AnalysisEventListener即可
     */
    @Test
    public void readByListen() {
        final List<String> lists = new ArrayList<String>();
        EasyExcelUtil.read("C:\\Users\\Draher\\Desktop\\sku_status_import_template.xlsx", new AnalysisEventListener() {
            @Override
            @SuppressWarnings("unchecked")
            public void invoke(Object o, AnalysisContext analysisContext) {
                JSONObject obj = new JSONObject();
                List<String> values = (List<String>) o;
                List<String> titles = Arrays.asList("plat", "wh", "account", "site", "spu", "sku", "onlineSku", "status", "reason");
                for (int i = 0; i < titles.size(); i++) {
                    obj.put(titles.get(i), values.get(i));
                }
                lists.add(obj.toJSONString());
            }

            @Override
            public void doAfterAllAnalysed(AnalysisContext analysisContext) {
                System.out.println(lists.size());
            }
        });
    }

    @Test
    public void writeByList() {
        List<List<Object>> list = new ArrayList<List<Object>>();
        for (int i = 0; i < 1000; i++) {
            List<Object> da = new ArrayList<Object>();
            da.add("字符串222222222222222222222222" + i);
            da.add(187837834L + i);
            da.add(2233 + i);
            da.add(2233.01 + i);
            da.add(2233.2f + i);
            da.add(new Date());
            da.add(new BigDecimal("3434343433554545555555" + i));
            da.add((short) i);
            list.add(da);
        }

        long start = System.currentTimeMillis();
        EasyExcelUtil.write(list, Arrays.asList("第一列", "第二列", "第三列"), "hello world", "C:\\Users\\Draher\\Desktop\\string.xlsx");
        long end = System.currentTimeMillis();
        System.out.println("run times：" + (end - start));
    }

    /**
     * 通过javaBean写入数据
     */
    @Test
    public void writeByBean() {
        long start = System.currentTimeMillis();
        List<SkuStatusTO> list = createJavaMode();
        EasyExcelUtil.writeByBean(list, SkuStatusTO.class, "hello world", "C:\\Users\\Draher\\Desktop\\writeModel.xlsx");
        long end = System.currentTimeMillis();
        System.out.println("总运行时间:" + ( end - start));
    }

    /**
     * 通过pi原先类似的方式写excel
     * 通过封装WriteInfo对象， WriteInfo只能通过构建器Build创建
     * @throws IllegalAccessException
     */
    @Test
    public void write() throws IllegalAccessException {
        EasyExcelUtil.write(new WriteInfo.Builder()
                .title(new String[]{"平台", "仓库", "账号", "站点", "spu", "sku", "在线sku", "在线状态", "原因"})
                .sheetName("hello")
                .contentTitle(new String[]{"plat", "wh", "account", "site", "spu", "sku", "onlineSku", "status", "reason"})
                .contentList(createDbMapData())
                .build(), "C:\\Users\\Draher\\Desktop\\writeInfo.xlsx");
    }

    /**
     * 写入添加回调，如可将cell值为fail的单元格置为红色等
     * @throws IllegalAccessException
     */
    @Test
    public void writeWithHandle() throws IllegalAccessException {
        WriteInfo writeInfo = new WriteInfo.Builder()
                .title(new String[]{"平台", "仓库", "账号", "站点", "spu", "sku", "在线sku", "在线状态", "原因"})
                .sheetName("hello")
                .contentTitle(new String[]{"plat", "wh", "account", "site", "spu", "sku", "onlineSku", "status", "reason"})
                .contentList(createDbMapData())
                .build();

        EasyExcelUtil.write(writeInfo, "C:\\Users\\Draher\\Desktop\\writeInfo.xlsx", new WriteHandler() {
            @Override
            public void sheet(int sheetNo, Sheet sheet) {

            }

            @Override
            public void row(int rowNum, Row row) {

            }

            @Override
            public void cell(int cellNum, Cell cell, Object cellValue) {
                Workbook workbook = cell.getSheet().getWorkbook();
                if (cellNum == 8 && cell.getRowIndex() > 0 && cellValue.equals("success1")) {
                    CellStyle  newCellStyle = workbook.createCellStyle();
                    newCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                    newCellStyle.setAlignment(HorizontalAlignment.CENTER);
                    newCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                    newCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    cell.setCellStyle(newCellStyle);
                }
            }
        });
    }


    private List<SkuStatusTO> createJavaMode() {
        List<SkuStatusTO> list = new ArrayList<SkuStatusTO>();
        for (int i = 0; i < 1000; i++) {
            SkuStatusTO skuStatusTO = new SkuStatusTO();
            skuStatusTO.setPlat("amazon" + i);
            skuStatusTO.setWh("fba" + i);
            skuStatusTO.setAccount("avidlove11111111111111111111111" + i);
            skuStatusTO.setAccountId(1);
            skuStatusTO.setSite("us" + i);
            skuStatusTO.setSpu("aaa" + i);
            skuStatusTO.setSku("bbb" + i);
            skuStatusTO.setOnlineSku("abab" + i);
            skuStatusTO.setStatus("ccc" + i);
            skuStatusTO.setReason("success" + i);
            skuStatusTO.setObject(new JSONObject(){{put("name", "测试");}});
            skuStatusTO.setDate(new Date());
            list.add(skuStatusTO);
        }

        return list;
    }

    private List<Map<String, Object>> createDbMapData() throws IllegalAccessException {
        return BeanConvertUtil.listBean2Map(createJavaMode());
    }

    /**
     * 用法类似fastJson的类注解
     *
     * @see com.sailvan.excel.annotation.ExcelProperty 注解
     *
     * index  数值对应列，默认值为9999
     * value  对应表头,表头可设置为单行表头，或者多行的复合表头
     * convertor 属性可实现类型转换，如accountId -> Account, 或 JSONObject -> jsonString提供自定义接口,实现TypeConvertor即可
     *           底层使用反射实现，使用可能会稍微影响一点程序性能
     *           因为MVC设计原则，Model层不依赖与任何业务逻辑，如果需要实现acccountId -> account的转换，这里涉及到DAO层操作（因为需要查询数据库当前的account），
     *           则需要在service层创建一个专门处理excel的model类（推荐静态内部类），才可以添加转换的TypeConvertor
     * format 对日期类型进行转换
     * ignore 可忽略无需解析的字段，一般配合ExcelProperty注解使用在类名的情况
     * order 字段排序，底层原理是重设index属性的变量，注意必须包含全部需要写入的字段，不然会抛出RunTimeException
     *
     * 如果未在类名上添加注解，则会按照有加注解的成员变量和index属性进行写入excel(仅处理有ExcelProperty注解的field)；
     * 如果在类名添加注解，则默认所有属性都将进行写入excel，未指定表头名称的（即value属性），则按字段名设置表头，index值为字段的默认排序。
     * orders属性会重排字段的index数值。
     *
     */
    @ExcelProperty(orders = {"wh", "plat", "account", "accountId", "site", "spu", "sku", "onlineSku", "status", "reason", "object", "date"})
    public static class SkuStatusTO extends BaseRowModel {
        @ExcelProperty(value = "平台")
        private String plat;
        @ExcelProperty(value = "仓库")
        private String wh;
        @ExcelProperty(value = "账号")
        private String account;
        @ExcelProperty(value = "账号Id", convertor = AccountTypeConvertor.class)
        private Integer accountId;
        @ExcelProperty(value = "站点")
        private String site;
        private String spu;
        private String sku;
        @ExcelProperty(value = "在线sku")
        private String onlineSku;
        @ExcelProperty(value = "在线状态")
        private String status;
        @ExcelProperty(value = "原因")
        private String reason;
        @ExcelProperty(value = "序列号", convertor = JsonTypeConvertor.class)
        private JSONObject object;
        @ExcelProperty(value = "时间", format = "yyyy-MM-dd")
        private Date date;

        public String getPlat() {
            return plat;
        }

        public void setPlat(String plat) {
            this.plat = plat;
        }

        public String getWh() {
            return wh;
        }

        public void setWh(String wh) {
            this.wh = wh;
        }

        public String getAccount() {
            return account;
        }

        public void setAccount(String account) {
            this.account = account;
        }

        public Integer getAccountId() {
            return accountId;
        }

        public void setAccountId(Integer accountId) {
            this.accountId = accountId;
        }

        public String getSite() {
            return site;
        }

        public void setSite(String site) {
            this.site = site;
        }

        public String getSpu() {
            return spu;
        }

        public void setSpu(String spu) {
            this.spu = spu;
        }

        public String getSku() {
            return sku;
        }

        public void setSku(String sku) {
            this.sku = sku;
        }

        public String getOnlineSku() {
            return onlineSku;
        }

        public void setOnlineSku(String onlineSku) {
            this.onlineSku = onlineSku;
        }

        public String getStatus() {
            return status;
        }

        public void setStatus(String status) {
            this.status = status;
        }

        public String getReason() {
            return reason;
        }

        public void setReason(String reason) {
            this.reason = reason;
        }

        public JSONObject getObject() {
            return object;
        }

        public void setObject(JSONObject object) {
            this.object = object;
        }

        public Date getDate() {
            return date;
        }

        public void setDate(Date date) {
            this.date = date;
        }
    }

    public static class BeanConvertUtil {
        public static <T> Map<String, Object> bean2Map(T bean) throws IllegalAccessException {
            Map<String, Object> map = new HashMap<String, Object>();
            Class<?> clazz = bean.getClass();
            // 获取所有字段（不包括父类）
            Field[] fields = clazz.getDeclaredFields();
            for (Field field : fields) {
                field.setAccessible(true);
                map.put(field.getName(), field.get(bean));
            }

            return map;
        }

        public static <T> List<Map<String, Object>> listBean2Map(List<T> beanList)
                throws IllegalAccessException {
            List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
            for (T t : beanList) {
                list.add(bean2Map(t));
            }

            return list;
        }
    }

}
