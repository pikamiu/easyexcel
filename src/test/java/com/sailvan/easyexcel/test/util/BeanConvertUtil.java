package com.sailvan.easyexcel.test.util;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


/**
 * <p>简要说明...</p>
 *
 * @author wujiaming
 * @version 1.0
 * @since 2018/8/2 15:37
 */
public class BeanConvertUtil {

    /**
     * bean 转 Map
     * @param bean 对象
     * @param <T> 对象类型
     * @return Map<String, Object>
     */
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
