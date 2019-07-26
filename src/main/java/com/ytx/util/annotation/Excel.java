package com.ytx.util.annotation;

import com.ytx.util.enums.ExcelType;

import java.lang.annotation.*;

/**
 * @author wangqi
 * @version 1.0
 * @className Excel
 * @description Excel对象枚举
 * @date 2019/7/15 18:28
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE})
@Documented
public @interface Excel {
    /**
     * 导出excel表名
     * @return
     */
    String value() default "";

    /**
     * 导出excel类型
     * @return
     */
    ExcelType type() default ExcelType.XLS;

    /**
     * 默认读取的sheet，如果为空，读取全部
     * @return
     */
    ExcelSheet[] sheet() default {};

}

