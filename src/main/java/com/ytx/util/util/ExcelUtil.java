package com.ytx.util.util;


import com.ytx.util.annotation.Excel;
import com.ytx.util.annotation.ExcelFieldChange;
import com.ytx.util.annotation.ExcelSheet;
import com.ytx.util.annotation.ExcelField;
import com.ytx.util.entity.ExcelParam;
import com.ytx.util.entity.SheetParam;
import com.ytx.util.enums.ExcelType;
import com.ytx.util.exception.ExcelException;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @author wangqi
 * @version 1.0
 * @className ExcelUtil
 * @description Excel工具类
 * @date 2019/7/8 11:16
 */
class ExcelUtil implements Serializable {

    private static final long serialVersionUID = 1L;

    /**
     * 默认sheet对象
     */
    protected static SheetParam initSheetParam;

    /**
     * 初始化默认sheet对象
     */
    static {
        initSheetParam = new SheetParam();
        initSheetParam.setTitleIndex(0).setStartIndex(1);
    }

    /**
     * 读取单元格值
     * @param cell 单元格
     * @param field 单元格字段
     * @return 单元格值
     */
    protected static Object getValue(Cell cell, Field field) {
        ExcelFieldChange[] fieldChanges = field.getAnnotation(ExcelField.class).fieldChange();
        Class clazz = field.getType();
        Object val = null;
        if (cell.getCellTypeEnum() == CellType.BOOLEAN) {
            val = cell.getBooleanCellValue();
        } else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
            double value = cell.getNumericCellValue();
            if (DateUtil.isCellDateFormatted(cell)) {
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                if (clazz == String.class) {
                    val = sdf.format(DateUtil.getJavaDate(value));
                } else if (clazz == int.class || clazz == Integer.class) {
                    val = DateUtil.getJavaDate(value).getTime() / 1000;
                } else if (clazz == long.class || clazz == Long.class) {
                    val = DateUtil.getJavaDate(value).getTime();
                } else {
                    val = DateUtil.getJavaDate(value);
                }
            } else {
                if (clazz == String.class) {
                    cell.setCellType(CellType.STRING);
                    val = cell.getStringCellValue();
                } else if (clazz == BigDecimal.class) {
                    val = new BigDecimal(value);
                } else if (clazz == long.class || clazz == Long.class) {
                    val = (long) value;
                } else if (clazz == Double.class || clazz == double.class) {
                    val = value;
                } else if (clazz == Float.class || clazz == float.class) {
                    val = (float) value;
                } else if (clazz == int.class || clazz == Integer.class) {
                    val = (int) value;
                } else if (clazz == Short.class || clazz == short.class) {
                    val = (short) value;
                } else {
                    val = value;
                }
            }
        } else if (cell.getCellTypeEnum() == CellType.STRING) {
            val = cell.getStringCellValue();
        }
        if(fieldChanges != null && fieldChanges.length > 0){
            for(ExcelFieldChange fieldChange:fieldChanges){
                if(val != null && fieldChange.value().equals(val.toString())){
                    String key = fieldChange.key();
                    if (clazz == BigDecimal.class) {
                        val = new BigDecimal(key);
                    } else if (clazz == long.class || clazz == Long.class) {
                        val = Long.valueOf(key);
                    } else if (clazz == Double.class || clazz == double.class) {
                        val = Double.valueOf(key);
                    } else if (clazz == Float.class|| clazz == float.class) {
                        val = Float.valueOf(key);
                    } else if (clazz == int.class || clazz == Integer.class) {
                        val = Integer.valueOf(key);
                    } else if (clazz == Short.class|| clazz == short.class) {
                        val = Short.valueOf(key);
                    }else if (clazz == Byte.class|| clazz == byte.class) {
                        val = Byte.valueOf(key);
                    }else if (clazz == Boolean.class|| clazz == boolean.class) {
                        val = Boolean.valueOf(key);
                    }
                }
            }
        }
        return val;
    }

    protected static <T> void getExcelParam(ExcelParam<T> excelParam,List<T> ...lists){
        ExcelUtil.getExcelAnnotation(excelParam);
        if(lists == null || lists.length==0){
            throw new ExcelException("导出数据不能为空");
        }
        SheetParam<T>[] sheetParams = excelParam.getSheetParams();
        if(sheetParams == null || sheetParams.length==0){
            sheetParams = new SheetParam[lists.length];
        }else{
            sheetParams = Arrays.copyOf(sheetParams,lists.length);
        }
        for (int i = 0;i<lists.length;i++){
            if(lists[i] == null){
                throw new ExcelException("导出数据不能为空");
            }
            SheetParam<T> sheetParam = sheetParams[i];
            if(sheetParam == null){
                sheetParam = new SheetParam<>();
                sheetParam.setSheetIndex(i).setSheetName("Sheet"+(i+1)).setTitleIndex(0).setStartIndex(1);
            }
            sheetParam.setList(lists[i]);
            sheetParams[i] = sheetParam;
        }
        excelParam.setSheetParams(sheetParams);
    }

    /**
     * 获取class对象所有的相关注解
     * 如果ExcelParam对象中存在该属性，则覆盖注解对象
     * @param excelParam excel参数对象
     */
    protected static<T> void getExcelAnnotation(ExcelParam excelParam) {
        Class<T> clazz = excelParam.getClazz();
        if(clazz ==null){
            throw new ExcelException("接收对象为空");
        }
        getFieldsMaps(excelParam);
        //获取对象注解Excel的数据，如果ExcelParam对象中存在该属性，则覆盖注解对象
        Excel excel = clazz.getAnnotation(Excel.class);
        if(excel == null) {
            return;
        }
        //获取文件名称
        if(StringUtils.isEmpty(excelParam.getFileName())&&!StringUtils.isEmpty(excel.value())){
            excelParam.setFileName(excel.value());
        }
        if(excelParam.getType() == null){
            excelParam.setType(excel.type());
        }
        if((excelParam.getSheetParams() == null || excelParam.getSheetParams().length == 0) && excel.sheet()!= null && excel.sheet().length != 0){
            SheetParam[] excelSheets = new SheetParam[excel.sheet().length];
            for(int i = 0;i < excel.sheet().length;i++){
                ExcelSheet excelSheet = excel.sheet()[i];
                SheetParam sheetParam = new SheetParam();
                //获取需要读取的sheet
                if(excelSheet.sheetIndex() != 0){
                    sheetParam.setSheetIndex(excelSheet.sheetIndex());
                }
                if(StringUtils.isNotEmpty(excelSheet.sheetName())){
                    sheetParam.setSheetName(excelSheet.sheetName());
                }
                //获取标题默认所在行数
                if(excelSheet.titleIndex() != 0){
                    sheetParam.setTitleIndex(excelSheet.titleIndex());
                }
                //获取开始读取的行数
                if(excelSheet.startIndex() != 1){
                    sheetParam.setStartIndex(excelSheet.startIndex());
                }
                //获取每次读取条数限制
                if(excelSheet.length() != 0){
                    sheetParam.setLength(excelSheet.length());
                }
                //是否采用兼容模式
                if(excelSheet.compatible()){
                    sheetParam.setCompatible(excelSheet.compatible());
                }
                excelSheets[i]=sheetParam;
            }
            excelParam.setSheetParams(excelSheets);
        }
    }

    /**
     * 获取所有支持导入导出的属性
     * @param excelParam excel参数对象
     * @param <T>
     */
    private static<T> void getFieldsMaps(ExcelParam excelParam) {
        Class<T> clazz = excelParam.getClazz();
        if(clazz ==null){
            throw new ExcelException("接收对象为空");
        }
        Map<String, Field> fieldsMap = excelParam.getFieldsMap();
        if(fieldsMap == null){
            fieldsMap = new HashMap();
            excelParam.setFieldsMap(fieldsMap);
        }
        // 获取所有支持导入导出的属性
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (excelField != null && StringUtils.isNotEmpty(excelField.value())) {
                ExcelFieldChange[] fieldChanges = excelField.fieldChange();
                if(fieldChanges != null && fieldChanges.length > 0){
                    for(ExcelFieldChange fieldChange:fieldChanges){
                        if(StringUtils.isEmpty(fieldChange.key()) || StringUtils.isEmpty(fieldChange.value())){
                            throw new ExcelException("excel和java对象转换设置为空");
                        }
                    }
                }
                if(!fieldsMap.containsKey(excelField.value()) || !fieldsMap.containsValue(field)){
                    fieldsMap.put(excelField.value(), field);
                }
            }
        }
    }
}