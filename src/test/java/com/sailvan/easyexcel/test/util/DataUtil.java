package com.sailvan.easyexcel.test.util;

import com.sailvan.easyexcel.test.model.WriteModel;
import com.sailvan.easyexcel.test.model.WriteModel2;
import com.sailvan.excel.metadata.Font;
import com.sailvan.excel.metadata.TableStyle;
import com.alibaba.fastjson.JSONObject;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

public class DataUtil {


    public static List<List<Object>> createTestListObject() {
        List<List<Object>> object = new ArrayList<List<Object>>();
        for (int i = 0; i < 1000; i++) {
            List<Object> da = new ArrayList<Object>();
            da.add("字符串"+i);
            da.add(Long.valueOf(187837834l+i));
            da.add(Integer.valueOf(2233+i));
            da.add(Double.valueOf(2233.00+i));
            da.add(Float.valueOf(2233.0f+i));
            da.add(new Date());
            da.add(new BigDecimal("3434343433554545"+i));
            da.add(Short.valueOf((short)i));
            object.add(da);
        }
        return object;
    }

    public static List<String> createTestJsonString() {
        List<String> object = new ArrayList<String>();
        for (int i = 0; i < 1000; i++) {
            WriteModel model1 = new WriteModel();
            model1.setP1("第一列，第行");
            model1.setP2("121212jjj");
            model1.setP3(33+i);
            model1.setP4(44);
            model1.setP5("555");
            model1.setP6(666.2f);
            model1.setP7(new BigDecimal("454545656343434"+i));
            model1.setP8(new Date());
            model1.setP9("llll9999>&&&&&6666^^^^");
            model1.setP10(1111.77+i);
            object.add(JSONObject.toJSONString(model1));
        }
        return object;
    }

    public static List<List<String>> createTestListStringHead(){
        //写sheet3  模型上没有注解，表头数据动态传入
        List<List<String>> head = new ArrayList<List<String>>();
        List<String> headCoulumn1 = new ArrayList<String>();
        List<String> headCoulumn2 = new ArrayList<String>();
        List<String> headCoulumn3 = new ArrayList<String>();
        List<String> headCoulumn4 = new ArrayList<String>();
        List<String> headCoulumn5 = new ArrayList<String>();

        headCoulumn1.add("第一列");headCoulumn1.add("第一列");headCoulumn1.add("第一列");
        headCoulumn2.add("第一列");headCoulumn2.add("第一列");headCoulumn2.add("第一列");

        headCoulumn3.add("第二列");headCoulumn3.add("第二列");headCoulumn3.add("第二列");
        headCoulumn4.add("第三列");headCoulumn4.add("第三列2");headCoulumn4.add("第三列2");
        headCoulumn5.add("第一列");headCoulumn5.add("第3列");headCoulumn5.add("第4列");

        head.add(headCoulumn1);
        head.add(headCoulumn2);
        head.add(headCoulumn3);
        head.add(headCoulumn4);
        head.add(headCoulumn5);
        return head;
    }

    public static List<WriteModel> createTestListJavaMode(){
        List<WriteModel> model1s = new ArrayList<WriteModel>();
        for (int i = 0; i <10000 ; i++) {
            WriteModel model1 = new WriteModel();
            model1.setP1("第一列，第行");
            model1.setP2("121212jjj");
            model1.setP3(33+i);
            model1.setP4(44);
            model1.setP5("555");
            model1.setP6(666.2f);
            model1.setP7(new BigDecimal("454545656343434"+i));
            model1.setP8(new Date());
            model1.setP9("llll9999>&&&&&6666^^^^");
            model1.setP10(1111.77+i);
            model1s.add(model1);
        }
        return model1s;
    }


    public static List<WriteModel2> createTestListJavaMode2(){
        List<WriteModel2> model1s = new ArrayList<WriteModel2>();
        for (int i = 0; i <10000 ; i++) {
            WriteModel2 model1 = new WriteModel2();
            model1.setP1("第一列，第行");
            model1.setP2("121212jjj");
            model1.setP3(33+i);
            model1.setP4(44);
            model1.setP5("555");
            model1.setP6(666.2f);
            model1.setP7(new BigDecimal("454545656343434"+i));
            model1.setP8(new Date());
            model1.setP9("llll9999>&&&&&6666^^^^");
            model1.setP10(1111.77+i);
            model1s.add(model1);
        }
        return model1s;
    }

    public static TableStyle createTableStyle() {
        TableStyle tableStyle = new TableStyle();
        Font headFont = new Font();
        headFont.setBold(true);
        headFont.setFontHeightInPoints((short)22);
        headFont.setFontName("楷体");
        tableStyle.setTableHeadFont(headFont);
        tableStyle.setTableHeadBackGroundColor(IndexedColors.BLUE);

        Font contentFont = new Font();
        contentFont.setBold(true);
        contentFont.setFontHeightInPoints((short)22);
        contentFont.setFontName("黑体");
        tableStyle.setTableContentFont(contentFont);
        tableStyle.setTableContentBackGroundColor(IndexedColors.GREEN);
        return tableStyle;
    }

    public static TableStyle createPITableStyle() {
        TableStyle tableStyle = new TableStyle();
        Font headFont = new Font();
        headFont.setBold(false);
        headFont.setFontName("宋体");
        headFont.setFontHeightInPoints((short) 11);
        tableStyle.setTableHeadFont(headFont);
        tableStyle.setTableHeadBackGroundColor(IndexedColors.PALE_BLUE);


        Font contentFont = new Font();
        contentFont.setBold(false);
        contentFont.setFontHeightInPoints((short)11);
        contentFont.setFontName("宋体");
        tableStyle.setTableContentFont(contentFont);
        tableStyle.setTableContentBackGroundColor(IndexedColors.WHITE1);

        return tableStyle;
    }

    public static List<Map<String, Object>> createDbMapData() throws IllegalAccessException {
        return BeanConvertUtil.listBean2Map(createTestListJavaMode2());
    }
}
