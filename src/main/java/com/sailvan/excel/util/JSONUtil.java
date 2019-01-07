package com.sailvan.excel.util;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.TypeReference;
import com.alibaba.fastjson.parser.Feature;

/**
 * <p>简要说明...</p>
 *
 * @author wujiaming
 * @version 1.0
 * @since 2018/12/28 16:48
 */
public class JSONUtil {

    public static List<String> jsonParse(String str) {
        LinkedHashMap<String, Object> rootStr = JSON.parseObject(str, new TypeReference<LinkedHashMap<String, Object>>() {}, Feature.OrderedField);
        return new ArrayList<String>(rootStr.keySet());
    }
}
