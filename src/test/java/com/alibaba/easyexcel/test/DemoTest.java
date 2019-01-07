package com.alibaba.easyexcel.test;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;

import com.alibaba.easyexcel.test.util.BeanConvertUtil;
import com.alibaba.excel.EasyExcelUtil;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.WriteInfo;
import com.alibaba.fastjson.JSONObject;

import jdk.nashorn.internal.objects.annotations.Getter;

/**
 * <p>简要说明...</p>
 *
 * @author wujiaming
 * @version 1.0
 * @since 2019/1/6 19:50
 */
public class DemoTest {

    @Test
    public void read() {
        List<List<String>> list = EasyExcelUtil.read("C:\\Users\\Draher\\Desktop\\sku_status_import_template.xlsx");
        System.out.println(list.size());
    }

    @Test
    public void readByModel() {
        List<SkuStatusTO> list = EasyExcelUtil.read("C:\\Users\\Draher\\Desktop\\sku_status_import_template.xlsx", SkuStatusTO.class);
        System.out.println(list.size());
    }

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
            da.add("字符串" + i);
            da.add(187837834L + i);
            da.add(2233 + i);
            da.add(2233.01 + i);
            da.add(2233.2f + i);
            da.add(new Date());
            da.add(new BigDecimal("3434343433554545" + i));
            da.add((short) i);
            list.add(da);
        }

        EasyExcelUtil.write(list, Arrays.asList("第一列", "第二列", "第三列"), "hello world", "C:\\Users\\Draher\\Desktop\\string.xlsx");
    }

    @Test
    public void writeByBean() {
        List<SkuStatusTO> list = createJavaMode();
        EasyExcelUtil.writeByBean(list, SkuStatusTO.class, "hello world", "C:\\Users\\Draher\\Desktop\\writeModel.xlsx");
    }

    @Test
    public void write() throws IllegalAccessException {
        EasyExcelUtil.write(new WriteInfo.Builder()
                .title(new String[]{"平台", "仓库", "账号", "站点", "spu", "sku", "在线sku", "在线状态", "原因"})
                .sheetName("hello")
                .contentTitle(new String[]{"plat", "wh", "account", "site", "spu", "sku", "onlineSku", "status", "reason"})
                .contentList(createDbMapData())
                .build(), "C:\\Users\\Draher\\Desktop\\writeInfo.xlsx");
    }

    private List<SkuStatusTO> createJavaMode() {
        List<SkuStatusTO> list = new ArrayList<SkuStatusTO>();
        for (int i = 0; i < 10; i++) {
            SkuStatusTO skuStatusTO = new SkuStatusTO();
            skuStatusTO.setPlat("amazon" + i);
            skuStatusTO.setWh("fba" + i);
            skuStatusTO.setAccount("avidlove" + i);
            skuStatusTO.setSite("us" + i);
            skuStatusTO.setSpu("aaa" + i);
            skuStatusTO.setSku("bbb" + i);
            skuStatusTO.setOnlineSku("abab" + i);
            skuStatusTO.setStatus("ccc" + i);
            skuStatusTO.setReason("test" + i);
            list.add(skuStatusTO);
        }

        return list;
    }

    private List<Map<String, Object>> createDbMapData() throws IllegalAccessException {
        return BeanConvertUtil.listBean2Map(createJavaMode());
    }

    @ExcelProperty(orders = {"wh", "plat", "account", "site", "spu", "sku", "onlineSku", "reason", "status"})
    private static class SkuStatusTO extends BaseRowModel {
        @ExcelProperty(value = "平台")
        private String plat;
        @ExcelProperty(value = "仓库")
        private String wh;
        @ExcelProperty(value = "账号")
        private String account;
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
