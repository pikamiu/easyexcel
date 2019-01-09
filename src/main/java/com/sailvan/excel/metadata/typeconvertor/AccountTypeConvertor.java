package com.sailvan.excel.metadata.typeconvertor;

import java.util.HashMap;
import java.util.Map;

/**
 * <p>简要说明...</p>
 *
 * @author wujiaming
 * @version 1.0
 * @since 2019/1/9 17:26
 */
public class AccountTypeConvertor implements TypeConvertor<Integer> {
    private static Map<String, Integer> map;
    private static Map<Integer, String> deMap;

    @Override
    public Integer reversal(String s) {
        if (map == null) {
            serializeInit();
        }
        return map.get(s);
    }

    @Override
    public String convert(Integer s) {
        if (deMap == null) {
            deserializeInit();
        }
        return deMap.get(s);
    }

    private void deserializeInit() {
        deMap = new HashMap<Integer, String>(){{
            put(1, "zeagoo889");
        }};
    }

    private void serializeInit() {
        map = new HashMap<String, Integer>(){{
            put("zeagoo889", 1);
        }};
    }

    @Override
    public void doAfter() {
        map = null;
        deMap = null;
    }
}
