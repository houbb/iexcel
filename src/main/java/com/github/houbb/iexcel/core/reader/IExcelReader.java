package com.github.houbb.iexcel.core.reader;

import java.util.List;
import java.util.Map;

/**
 * @author binbin.hou
 * @date 2018/11/15 19:57
 */
public interface IExcelReader {

    <T> List<T> readAll(Class<T> tClass);

    List<Map<String, Object>> readAll();


}
