package com.ytx.util.annotation;

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
     * 默认读取的sheet，如果为空，读取全部
     * @return
     */
    ExcelSheet[] sheet() default {};

}

