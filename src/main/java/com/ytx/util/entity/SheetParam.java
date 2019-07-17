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
     * 标题默认所在行数
     * @return
     */
    Integer titleIndex;

    /**
     * 数据默认开始读取插入行数
     * @return
     */
    Integer startIndex;

    /**
     * 每次读取条数限制
     * @return
     */
    Integer length;

    /**
     * 	存储字段和Excel列的对应关系
     * 	key是表头名称
     * 	value是字段Field
     */
    Map<Integer, Field> titleMap;
}
