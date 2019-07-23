package com.ytx.util.annotation;

import java.lang.annotation.*;

/**
 * @author wangqi
 * @version 1.0
 * @className ExcelFieldChange
 * @description excel和java对象转换
 * @date 2019/7/22 18:58
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
@Documented
public @interface ExcelFieldChange {

    String key() ;

    String value();
}
