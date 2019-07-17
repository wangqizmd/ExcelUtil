package com.ytx.util.annotation;

import java.lang.annotation.*;

/**
 * @author wangqi
 * @version 1.0
 * @className ExcelSheet
 * @description TODO
 * @date 2019/7/16 15:01
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE})
@Documented
public @interface ExcelSheet {
    /**
     * 读取的sheetIndex
     * @return
     */
    int sheetIndex();
    /**
     * 标题默认所在行数
     * @return
     */
    int titleIndex() default 0;

    /**
     * 数据默认开始读取插入行数
     * @return
     */
    int startIndex() default 1;

    /**
     * 每次读取条数限制
     * @return
     */
    int length() ;
}
