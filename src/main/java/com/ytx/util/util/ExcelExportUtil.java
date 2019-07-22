package com.ytx.util.util;

import com.ytx.util.annotation.ExcelField;
import com.ytx.util.enums.ExcelType;
import com.ytx.util.exception.ExcelException;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.util.List;

/**
 * @author wangqi
 * @version 1.0
 * @className ExcelExportUtil
 * @description TODO
 * @date 2019/7/22 11:47
 */
public class ExcelExportUtil {
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
