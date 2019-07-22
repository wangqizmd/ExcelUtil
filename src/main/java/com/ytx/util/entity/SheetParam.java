package com.ytx.util.entity;

import lombok.Data;

import java.io.Serializable;
import java.lang.reflect.Field;
import java.util.Map;

/**
 * @author wangqi
 * @version 1.0
 * @className ExcelSheet
 * @description TODO
 * @date 2019/7/16 14:49
 */
@Data
public class SheetParam implements Serializable {
    private static final long serialVersionUID = 1L;

    /**
     * 读取的sheet index
     */
    int sheetIndex;

    /**
     * 读取的sheet name
     */
    String sheetName;

    /**
     * 标题默认所在行数
     */
    Integer titleIndex;

    /**
     * 数据默认开始读取/插入行数
     */
    Integer startIndex;

    /**
     * 每次读取/插入条数限制
     */
    Integer length;

    /**
     * 如果条数超过读取/插入条数限制，是否采用兼容模式，默认不采用
     * 该模式下超过读取/插入条数限制，不会抛出异常，只会读取/插入最大限制数据量
     */
    boolean compatible;

    /**
     * 	存储字段和Excel列的对应关系
     * 	key是表头名称
     * 	value是字段Field
     */
    Map<Integer, Field> titleMap;
}
