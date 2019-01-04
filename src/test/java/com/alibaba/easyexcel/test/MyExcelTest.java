package com.alibaba.easyexcel.test;

import static com.alibaba.easyexcel.test.util.DataUtil.*;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.junit.Test;

import com.alibaba.easyexcel.test.model.ReadModel;
import com.alibaba.easyexcel.test.model.WriteModel2;
import com.alibaba.excel.EasyExcelUtil;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Sheet.SheetBuilder;
import com.alibaba.excel.metadata.WriteInfo;

/**
 * <p>简要说明...</p>
 *
 * @author wujiaming
 * @version 1.0
 * @since 2018/12/27 15:15
 */
public class MyExcelTest {

    @Test
    public void simpleReadListStringV2007() throws IOException {
        List data = EasyExcelUtil.read("C:\\Users\\Draher\\Desktop\\sku_status_import_template.xlsx");
        print(data);
    }

    @Test
    public void simpleReadJavaModelV2007() throws IOException {
        List data = EasyExcelUtil.read("C:\\Users\\Draher\\Desktop\\11.xlsx", ReadModel.class);
        print(data);
    }

    @Test
    public void readDesign(){
        final List<List<String>> data = new ArrayList<List<String>>();
        EasyExcelUtil.read("C:\\Users\\Draher\\Desktop\\sku_status_import_template.xlsx",  new AnalysisEventListener() {
            @Override
            public void invoke(Object object, AnalysisContext context) {
                data.add((List<String>) object);
            }

            @Override
            public void doAfterAllAnalysed(AnalysisContext context) {

            }
        });
        System.out.println(data);
    }

    @Test
    public void writeByBean() {
        EasyExcelUtil.writeByBean(createTestListJavaMode2(), WriteModel2.class, "ah", "C:\\Users\\Draher\\Desktop\\2007model.xlsx");
    }

    @Test
    public void writeByBeanWithBuild() {
        EasyExcelUtil.writeByBean(createTestListJavaMode2(), new SheetBuilder()
                        .clazz(WriteModel2.class)
                        .autoWidth(true)
                        .sheetName("hello world")
                        .sheetNo(2)
                        .build(),
                "C:\\Users\\Draher\\Desktop\\2008model.xlsx");
    }

    @Test
    public void writeByStringWithMultiHead() {
        EasyExcelUtil.writeWithMultiHead(createTestListObject(), createTestListStringHead(), "C:\\Users\\Draher\\Desktop\\2009model.xlsx");
    }

    @Test
    public void writeByStringWithSingleHead() {
        EasyExcelUtil.write(createTestListObject(), Arrays.asList("第一列", "第二列", "第三列"), "C:\\Users\\Draher\\Desktop\\string.xlsx");
    }

    @Test
    public void writeByJson() {
        EasyExcelUtil.writeByJson(createTestJsonString(), "C:\\Users\\Draher\\Desktop\\json.xlsx");
    }

    @Test
    public void writeByJsonDIY() {
        EasyExcelUtil.writeByJson(createTestJsonString(), new SheetBuilder()
                .sheetName("hello world")
                .sheetNo(1)
                .headLineMun(1)
                .build(),"C:\\Users\\Draher\\Desktop\\json.xlsx");
    }

    @Test
    public void write() throws IllegalAccessException {
        Sheet mapSheet = new SheetBuilder()
                .sheetNo(4)
                .sheetName("hello hadoop")
                .contentTitle(Arrays.asList("p5", "p2", "p3"))
                .singleHead(Arrays.asList("账号", "平台", "站点"))
                .build();

        EasyExcelUtil.writeStream("C:\\Users\\Draher\\Desktop\\multiSheet.xlsx")
                .writeBean(createTestListJavaMode2(), new SheetBuilder().sheetNo(1).sheetName("hello java").clazz(WriteModel2.class).build())
                .writeObject(createTestListObject(), new SheetBuilder().sheetNo(2).sheetName("hello scala").head(createTestListStringHead()).build())
                .writeJson(createTestJsonString(), new SheetBuilder().sheetNo(3).sheetName("hello spark").build())
                .writeMap(createDbMapData(), mapSheet)
                .finish();
    }

    @Test
    public void writeMap() throws IllegalAccessException {
        final WriteInfo writeInfo = new WriteInfo.WriteInfoBuild()
                .sheetName("write test")
                .title(new String[]{"p1", "p2", "p3"})
                .contentTitle(new String[]{"p1", "p2", "p3"})
                .contentList(createDbMapData())
                .build();
        final WriteInfo writeInfo2 = new WriteInfo.WriteInfoBuild()
                .sheetName("write test2")
                .title(new String[]{"账号", "站点", "平台"})
                .contentTitle(new String[]{"p1", "p2", "p3"})
                .contentList(createDbMapData())
                .build();
        EasyExcelUtil.write(new ArrayList<WriteInfo>() {{
            add(writeInfo);
            add(writeInfo2);
        }}, "C:\\Users\\Draher\\Desktop\\map.xlsx");
    }


    public void print(List<Object> datas) {
        int i = 0;
        for (Object ob : datas) {
            System.out.println(i++);
            System.out.println(ob);
        }
    }
}
