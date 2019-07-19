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
     */
    int sheetIndex();
    /**
     * 标题默认所在行数
     */
    int titleIndex() default 0;

    /**
     * 数据默认开始读取插入行数
     */
    int startIndex() default 1;

    /**
     * 每次读取/插入条数限制
     */
    int length() ;

    /**
     * 如果条数超过读取/插入条数限制，是否采用兼容模式，，默认不采用
     * 如果为true,不会抛出异常，只会读取/插入最大限制数据量
     */
    boolean compatible() default false;
}
