package com.alibaba.excel.util;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * <p>简要说明...</p>
 *
 * @author wujiaming
 * @version 1.0
 * @since 2018/12/28 17:04
 */
public class ListUtil {
    public static List<List<String>> listWarp(List<String> list) {
        if (list == null) {
            return null;
        }
        List<List<String>> l = new ArrayList<List<String>>();
        for (String s : list) {
            l.add(Collections.singletonList(s));
        }
        return l;
    }
}
