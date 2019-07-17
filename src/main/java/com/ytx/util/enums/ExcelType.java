package com.ytx.util.enums;

import com.ytx.util.exception.ExcelException;

/**
 * @author wangqi
 * @version 1.0
 * @className ExcelType
 * @description EXCEL文件类型
 * @date 2019/7/15 17:47
 */

public enum ExcelType {

    /**
     * 03版Excel
     */
    XLS(1, "xls"),
    /**
     * 07版Excel
     */
    XLSX(2, "xlsx");

    private Integer key;
    private String value;

    ExcelType(Integer key, String value) {
        this.key = key;
        this.value = value;
    }

    public Integer getKey() {
        return key;
    }

    public void setKey(Integer key) {
        this.key = key;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public static ExcelType getValue(Integer key) {
        if(key == null){
            throw new ExcelException("请选择excel的文件类型");
        }
        ExcelType[] excelTypes = values();
        int length = excelTypes.length;

        for(int i = 0; i < length; i++) {
            ExcelType excelType = excelTypes[i];
            if (excelType.key.equals(key)) {
                return excelType;
            }
        }
        throw new ExcelException("请选择excel的文件类型");
    }
}
