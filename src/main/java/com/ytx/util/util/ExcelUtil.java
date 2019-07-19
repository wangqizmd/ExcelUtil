package com.ytx.util.util;


import com.ytx.util.annotation.Excel;
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
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author wangqi
 * @version 1.0
 * @className ExcelUtil
 * @description Excel工具类
 * @date 2019/7/8 11:16
 */
public class ExcelUtil implements Serializable {

    private static final long serialVersionUID = 1L;

    /**
     * 读取单元格值
     * @param cell 单元格
     * @param clazz 单元格数据类型
     * @return 单元格值
     */
    public static Object getValue(Cell cell, Class clazz) {
        Object val = null;
        if (cell.getCellTypeEnum() == CellType.BOOLEAN) {
            val = cell.getBooleanCellValue();
        } else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                if (clazz == String.class) {
                    val = sdf.format(DateUtil.getJavaDate(cell.getNumericCellValue()));
                } else if (clazz == int.class || clazz == Integer.class) {
                    val = DateUtil.getJavaDate(cell.getNumericCellValue()).getTime() / 1000;
                } else if (clazz == long.class || clazz == Long.class) {
                    val = DateUtil.getJavaDate(cell.getNumericCellValue()).getTime();
                } else {
                    val = DateUtil.getJavaDate(cell.getNumericCellValue());
                }
            } else {
                if (clazz == String.class) {
                    cell.setCellType(CellType.STRING);
                    val = cell.getStringCellValue();
                } else if (clazz == BigDecimal.class) {
                    val = new BigDecimal(cell.getNumericCellValue());
                } else if (clazz == long.class || clazz == Long.class) {
                    val = (long) cell.getNumericCellValue();
                } else if (clazz == Double.class) {
                    val = cell.getNumericCellValue();
                } else if (clazz == Float.class) {
                    val = (float) cell.getNumericCellValue();
                } else if (clazz == int.class || clazz == Integer.class) {
                    val = (int) cell.getNumericCellValue();
                } else if (clazz == Short.class) {
                    val = (short) cell.getNumericCellValue();
                } else {
                    val = cell.getNumericCellValue();
                }
            }
        } else if (cell.getCellTypeEnum() == CellType.STRING) {
            val = cell.getStringCellValue();
        }
        return val;
    }

    /**
     * 获取class对象所有的相关注解
     * 如果ExcelParam对象中存在该属性，则覆盖注解对象
     * @param excelParam excel参数对象
     */
    public static<T> void getExcelAnnotation(ExcelParam excelParam) {
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
        if(StringUtils.isEmpty(excelParam.getFileName())&&StringUtils.isEmpty(excel.value())){
            excelParam.setFileName(excel.value());
        }
        if((excelParam.getSheetParams() == null || excelParam.getSheetParams().length == 0) && excel.sheet()!= null && excel.sheet().length != 0){
            return;
        }
        SheetParam[] excelSheets = new SheetParam[excel.sheet().length];
        for(int i = 0;i < excel.sheet().length;i++){
            ExcelSheet excelSheet = excel.sheet()[i];
            SheetParam sheetParam = new SheetParam();
            //获取需要读取的sheet
            if(excelSheet.sheetIndex() != 0){
                sheetParam.setSheetIndex(excelSheet.sheetIndex());
            }
            //获取标题默认所在行数
            if(excelSheet.titleIndex() == 0){
                sheetParam.setTitleIndex(excelSheet.titleIndex());
            }
            //获取开始读取的行数
            if(excelSheet.startIndex() == 0){
                sheetParam.setStartIndex(excelSheet.startIndex());
            }
            //获取每次读取条数限制
            if(excelSheet.length() != 0){
                sheetParam.setLength(excelSheet.length());
            }
            excelSheets[i]=sheetParam;
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
            if (excelField != null && excelField.title() != null) {
                if(!fieldsMap.containsKey(excelField.title()) || !fieldsMap.containsValue(field)){
                    fieldsMap.put(excelField.title(), field);
                }
            }
        }
    }


    /**
     * 导出excel
     * @param list
     * @param clazz
     * @param response
     * @param fileName
     * @param titleIndex
     * @param <T>
     * @return
     */
    public static <T> void exportExcel(List<T> list, Class<T> clazz,
                                          HttpServletResponse response, String fileName, ExcelType excelType, Integer... titleIndex) {
        Workbook wb = exportWorkbook(list,clazz,excelType ,titleIndex);
        ServletOutputStream out = null;
        try {
            response.reset();
            response.setCharacterEncoding("UTF-8");
            response.setContentType("application/vnd.ms-excel");
            response.setHeader("Content-disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
            response.setHeader("Expires", "0");
            response.setHeader("Cache-Control", "must-revalidate, post-check=0,pre-check=0");
            response.setHeader("Pragma", "private");
            out = response.getOutputStream();
            wb.write(out);
            out.flush();
        } catch (Exception e) {
            throw new ExcelException("导出数据至excel文件失败");
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                }
            }
        }
    }

    /**
     * 创建HSSFWorkbook
     * @param list
     * @param clazz
     * @param titleIndex
     * @param <T>
     * @return
     */
    public static <T> Workbook exportWorkbook(List<T> list, Class<T> clazz, ExcelType excelType, Integer... titleIndex){
        if (excelType==null){
            throw new ExcelException("请选择excel的文件类型");
        }
        Workbook wb = null;
        if(ExcelType.XLS.equals(excelType)){
            // 创建HSSFWorkbook对象(excel的文档对象)
            wb = new HSSFWorkbook();
        }else if (ExcelType.XLSX.equals(excelType)){
            // 创建XSSFWorkbook对象(excel的文档对象)
            wb = new XSSFWorkbook();
        }
        // 建立新的sheet对象（excel的表单）
        Sheet sheet = wb.createSheet("sheet1");
        // 在sheet里创建第一行为表头
        setTitle(clazz, sheet,titleIndex);
        // 在sheet里创建表头下的数据
        if(CollectionUtils.isEmpty(list)){
            return wb;
        }
        int index = 0;
        if (titleIndex != null && titleIndex.length > 0){
            if(titleIndex[0]<=0){
                throw new ExcelException("表头行序号设置有误，应该大于0");
            }
            index = titleIndex[0]-1;
        }
        for (int i = 0; i < list.size(); i++) {
            setRow(list.get(i), clazz, sheet, index + i + 1);
        }
        return wb;
    }

    /**
     * 设置行的值
     * @param t
     * @param clazz
     * @param sheet
     * @param index
     * @param <T>
     */
    public static <T> void setRow(T t, Class<T> clazz, Sheet sheet, int index) {
        Row row = sheet.createRow(index);
        // 创建单元格并设置单元格内容
        Field[] fields = clazz.getDeclaredFields();
        int i =0;
        // 根据注解Csv写入表头信息
        for (Field field : fields) {
            field.setAccessible(true);
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (excelField != null && !excelField.ignore()) {
                try {
                    Cell cell = row.createCell(i++);
                    Object fieldData = field.get(t);
                    if(fieldData!=null){
                        cell.setCellValue(fieldData.toString());
                    }
                } catch (IllegalAccessException e) {
                    throw new ExcelException("读写字段属性"+field.getAnnotation(ExcelField.class).title()+"失败："+e);
                }
            }
        }
    }

    /**
     * 设置表头
     * @param clazz
     * @param sheet
     * @param titleIndex
     * @param <T>
     */
    public static <T> void setTitle(Class<T> clazz, Sheet sheet, Integer... titleIndex) {
        Row title = null;
        if (titleIndex != null && titleIndex.length > 0){
            if(titleIndex[0]<=0){
                throw new ExcelException("表头行序号设置有误，应该大于0");
            }
            title = sheet.createRow(titleIndex[0]-1);
        }else{
            title = sheet.createRow(0);
        }
        // 创建单元格并设置单元格内容
        Field[] fields = clazz.getDeclaredFields();
        int i=0;
        // 根据注解写入表头信息
        for (Field field : fields) {
            field.setAccessible(true);
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (excelField != null && !excelField.ignore()) {
                title.createCell(i++).setCellValue(excelField.title());
            }
        }
    }

}