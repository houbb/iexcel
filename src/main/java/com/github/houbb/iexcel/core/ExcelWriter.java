package com.github.houbb.iexcel.core;

import com.github.houbb.iexcel.metadata.Sheet;

import java.util.List;

/**
 * @author binbin.hou
 * @date 2018/11/1 10:57
 */
public interface ExcelWriter {

    void write(List<?> list, final Sheet sheet);

    void finish();

}
