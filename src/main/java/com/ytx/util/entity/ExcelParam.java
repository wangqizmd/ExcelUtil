package com.ytx.util.entity;

import com.ytx.util.enums.ExcelType;
import lombok.Data;
import lombok.experimental.Accessors;

import java.io.Serializable;
import java.lang.reflect.Field;
import java.util.Map;

/**
 * @author wangqi
 * @version 1.0
 * @className ExcelParam
 * @description Excel参数对象
 * @date 2019/7/16 11:57
 */
@Data
@Accessors(chain = true)
public class ExcelParam<T> implements Serializable {
    private static final long serialVersionUID = 1L;
    /**
     * 文件名称
     */
    String fileName;

    /**
     * 导出文件类型
     */
    ExcelType type;
    /**
     *导入/导出类
     */
    Class<T> clazz;


    /**
     * 当前正在读取/写入的sheet
     */
    SheetParam<T> sheetParam;
    /**
     * 需要读取/写入的sheet列表，如果为空，读取全部/写入第一个
     */
    SheetParam<T>[] sheetParams;
    /**
     * 	存储字段和表头的对应关系
     * 	key是表头名称
     * 	value是字段Field
     */
    Map<String, Field> fieldsMap;

}
