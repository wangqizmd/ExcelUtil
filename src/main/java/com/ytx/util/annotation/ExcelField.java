package com.ytx.util.annotation;

import org.springframework.core.annotation.AliasFor;

import java.lang.annotation.*;

/**
 * @author wangqi
 * @version 1.0
 * @className ExcelField
 * @description Excel对象字段枚举
 * @date 2019/7/16 11:26
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
@Documented
public @interface ExcelField {
    /**
     * 字段标题
     * @return
     */
    String value() ;

    /**
     * excel和java对象转换，比如1-男，2-女
     * @return
     */
    ExcelFieldChange[] fieldChange() default {};
    /**
     * 默认是否不为空
     * @return
     */
    boolean notNull() default true;

    /**
     * 导出默认是否忽略字段
     * @return
     */
    boolean ignore() default false;
}
