package com.github.houbb.iexcel.sax.handler;

import java.util.List;

/**
 * sax 单行信息处理接口
 * @author binbin.hou
 * @date 2018/11/16 13:57
 */
public interface SaxRowHandler {

    /**
     * 处理一行数据
     * @param context 上下文信息
     * @return 泛型类型
     */
    <T> T handle(SaxRowHandlerContext<T> context);

}
