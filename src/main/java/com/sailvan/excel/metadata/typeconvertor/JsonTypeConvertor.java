package com.sailvan.excel.metadata.typeconvertor;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;

/**
 * <p>简要说明...</p>
 *
 * @author wujiaming
 * @version 1.0
 * @since 2019/1/9 10:46
 */
public class JsonTypeConvertor implements TypeConvertor<JSONObject> {
    @Override
    public JSONObject reversal(String s) {
        return JSON.parseObject(s);
    }

    @Override
    public String convert(JSONObject s) {
        return s.toJSONString();
    }

    @Override
    public void doAfter() {

    }
}
