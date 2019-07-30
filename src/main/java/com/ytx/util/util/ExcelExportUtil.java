package com.ytx.util.util;

import com.ytx.util.annotation.ExcelField;
import com.ytx.util.annotation.ExcelFieldChange;
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
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author wangqi
 * @version 1.0
 * @className ExcelExportUtil
 * @description TODO
 * @date 2019/7/22 11:47
 */
public class ExcelExportUtil {


    public static <T> void exportExcel(Class<T> clazz, List<T>... list) throws IOException {
        exportExcel(clazz.getResource("/").getPath(),null,clazz,null,list);
    }

    public static <T> void exportExcel(String outFilePath,String fileName, Class<T> clazz, ExcelType excelType, List<T>... list) throws IOException {
        if(list == null || list.length==0){
            throw new ExcelException("导出数据不能为空");
        }
        ExcelParam<T> excelParam = new ExcelParam();
        excelParam.setClazz(clazz).setType(excelType).setFileName(fileName);
        ExcelUtil.getExcelParam(excelParam,list);
        Workbook wb = exportWorkbook(excelParam);
        FileOutputStream out = new FileOutputStream(new File(outFilePath,excelParam.getFileName()));
        wb.write(out);
        out.close();
        System.out.println(outFilePath+excelParam.getFileName());
    }

    public static <T> void exportExcel(Class<T> clazz, SheetParam<T> ...sheetParams) throws IOException {
        exportExcel( clazz.getResource("/").getPath(),null,clazz,null,sheetParams);
    }

    public static <T> void exportExcel(String outFilePath,String fileName, Class<T> clazz, ExcelType excelType, SheetParam<T> ...sheetParams) throws IOException {
        if(sheetParams == null || sheetParams.length==0){
            throw new ExcelException("导出数据不能为空");
        }
        ExcelParam<T> excelParam = new ExcelParam();
        excelParam.setClazz(clazz).setType(excelType).setSheetParams(sheetParams).setFileName(fileName);
        ExcelUtil.getExcelAnnotation(excelParam);
        Workbook wb = exportWorkbook(excelParam);
        FileOutputStream out = new FileOutputStream(new File(outFilePath,excelParam.getFileName()));
        wb.write(out);
        out.close();
        System.out.println(outFilePath+excelParam.getFileName());
    }

    public static <T> void exportExcel(HttpServletResponse response, Class<T> clazz, List<T>... list) {
        exportExcel(response,null,clazz,null,list);
    }

    public static <T> void exportExcel(HttpServletResponse response, String fileName, Class<T> clazz, ExcelType excelType, List<T>... list) {
        if(list == null || list.length==0){
            throw new ExcelException("导出数据不能为空");
        }
        ExcelParam<T> excelParam = new ExcelParam();
        excelParam.setClazz(clazz).setType(excelType).setFileName(fileName);
        ExcelUtil.getExcelParam(excelParam,list);
        Workbook wb = exportWorkbook(excelParam);
        exportExcel(response, excelParam, wb);
    }

    public static <T> void exportExcel(HttpServletResponse response, Class<T> clazz, SheetParam<T> ...sheetParams) {
        exportExcel(response,null,clazz,null,sheetParams);
    }

    public static <T> void exportExcel(HttpServletResponse response, String fileName, Class<T> clazz, ExcelType excelType, SheetParam<T> ...sheetParams) {
        if(sheetParams == null || sheetParams.length==0){
            throw new ExcelException("导出数据不能为空");
        }
        ExcelParam<T> excelParam = new ExcelParam();
        excelParam.setClazz(clazz).setType(excelType).setSheetParams(sheetParams).setFileName(fileName);
        ExcelUtil.getExcelAnnotation(excelParam);
        Workbook wb = exportWorkbook(excelParam);
        exportExcel(response, excelParam, wb);
    }

    private static <T> void exportExcel(HttpServletResponse response, ExcelParam<T> excelParam, Workbook wb) {
        ServletOutputStream out = null;
        try {
            response.reset();
            response.setCharacterEncoding("UTF-8");
            response.setContentType("application/vnd.ms-excel");
            response.setHeader("Content-disposition", "attachment;filename=" + URLEncoder.encode(excelParam.getFileName(), "UTF-8"));
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

    public static <T> Workbook exportWorkbook(Class<T> clazz,ExcelType excelType, SheetParam<T> ...sheetParams){
        if(sheetParams == null || sheetParams.length==0){
            throw new ExcelException("导出数据不能为空");
        }
        ExcelParam<T> excelParam = new ExcelParam();
        excelParam.setClazz(clazz).setType(excelType).setSheetParams(sheetParams);
        ExcelUtil.getExcelAnnotation(excelParam);
        return exportWorkbook(excelParam);
    }

    public static <T> Workbook exportWorkbook(Class<T> clazz,ExcelType excelType, List<T> ...list){
        if(list == null || list.length==0){
            throw new ExcelException("导出数据不能为空");
        }
        ExcelParam<T> excelParam = new ExcelParam();
        excelParam.setClazz(clazz).setType(excelType);
        ExcelUtil.getExcelParam(excelParam,list);
        return exportWorkbook(excelParam);
    }

    private static <T> Workbook exportWorkbook(ExcelParam<T> excelParam){
        SheetParam<T> []sheetParams = excelParam.getSheetParams();
        if (sheetParams==null || sheetParams.length==0){
            throw new ExcelException("导出数据不能为空");
        }
        Workbook wb = null;
        if(ExcelType.XLS.equals(excelParam.getType())){
            wb = new HSSFWorkbook();
        }else if (ExcelType.XLSX.equals(excelParam.getType())){
            wb = new XSSFWorkbook();
        }
        // 声明样式
        CellStyle style = wb.createCellStyle();
        // 居中显示
        style.setAlignment(HorizontalAlignment.CENTER);
        for (int i = 0;i<sheetParams.length;i++){
            SheetParam<T> sheetParam = sheetParams[i];
            if(sheetParam==null ){
                throw new ExcelException("导出配置不能为空");
            }
            Sheet sheet = wb.createSheet(StringUtils.isNotEmpty(sheetParam.getSheetName())?sheetParam.getSheetName():("Sheet"+(i+1)));
            setTitle(sheet,excelParam.getFieldsMap(),sheetParam);
            exportSheet(sheet,sheetParam);
        }
        return wb;
    }

    private static <T> void exportSheet(Sheet sheet,SheetParam<T> sheetParam) {
        List<T> list = sheetParam.getList();
        if(CollectionUtils.isEmpty(list)){
            throw new ExcelException("导出数据不能为空");
        }
        //如果存在条数限制，判断条数，以及处理策略
        Integer length = sheetParam.getLength();
        if(length != null ){
            if(list.size() > length){
                if(!sheetParam.isCompatible()){
                    throw new ExcelException("sheet:"+sheet.getSheetName()+"的插入条数超过最大条数限制"+length);
                }else{
                    list = list.subList(0,length);
                }
            }
        }
        //判断开始读取行数，默认是第一行，如果startIndex存在，则为startIndex，如果startIndex不存在，titleIndex存在，则为titleIndex+1行
        int startIndex = 1;
        if(sheetParam.getStartIndex()!=null && sheetParam.getStartIndex()!= 0){
            startIndex = sheetParam.getStartIndex();
        }else if(sheetParam.getTitleIndex()!=null && sheetParam.getTitleIndex()!= 0){
            startIndex = sheetParam.getTitleIndex() + 1;
        }
        for (int i=0;i<list.size();i++){
            try{
                setRow(sheet.createRow(i+startIndex),sheetParam.getTitleMap() ,list.get(i));
            }catch (ExcelException e){
                throw new ExcelException("sheet:"+sheet.getSheetName()+"第" + ( i+startIndex )+"行导出失败："+e.getMessage());
            }

        }
    }

    private static <T> void setRow(Row row,Map<Integer, Field> titleMap ,T t) {
        if(MapUtils.isEmpty(titleMap)){
            throw new ExcelException("无法匹配表头！");
        }
        for (int i = 0;i<titleMap.keySet().size();i++){
            setCell(row.createCell(i),titleMap.get(i),t);
        }

    }

    private static <T> void setCell(Cell cell,Field field, T t) {

        Object fieldData = null;
        try {
            field.setAccessible(true);
            fieldData = field.get(t);
            if (field.getAnnotation(ExcelField.class).notNull() && (fieldData == null || ("").equals(fieldData.toString().trim()))) {
                throw new ExcelException(field.getAnnotation(ExcelField.class).value() + "不能为空");
            }
            ExcelFieldChange[] fieldChanges = field.getAnnotation(ExcelField.class).fieldChange();
            if(fieldData!=null){
                if(fieldChanges != null && fieldChanges.length > 0) {
                    boolean flag = false;
                    for (ExcelFieldChange fieldChange : fieldChanges) {
                        if (fieldChange.key().equals(fieldData.toString())) {
                            flag = true;
                            fieldData = fieldChange.value();
                            break;
                        }
                    }
                    if (!flag) {
                        throw new ExcelException(field.getAnnotation(ExcelField.class).value() + "数据转换有误");
                    }
                }
                cell.setCellValue(fieldData.toString());
            }
        } catch (IllegalAccessException e) {
            throw new ExcelException("导入字段属性"+field.getAnnotation(ExcelField.class).value()+"失败："+e);
        }

    }

    private static <T> void setTitle(Sheet sheet,Map<String, Field> fieldsMap,SheetParam<T> sheetParam) {
        Row title = null;
        int titleIndex = 0;
        if(sheetParam.getTitleIndex()!=null && sheetParam.getTitleIndex()!= 0){
            if(sheetParam.getTitleIndex() < 0){
                throw new ExcelException("表头行序号设置有误，应该大于等于0");
            }
            titleIndex = sheetParam.getTitleIndex();
        }
        title = sheet.createRow(titleIndex);
        int i = 0;
        Map<Integer, Field> map = new HashMap<>();
        for (String key:fieldsMap.keySet()) {
            Field field = fieldsMap.get(key);
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (excelField != null && !excelField.ignore()) {
                title.createCell(i).setCellValue(key);
                map.put(i,field);
                i++;
            }
        }
        if(MapUtils.isEmpty(map)){
            throw new ExcelException("无法匹配表头！");
        }
        sheetParam.setTitleMap(map);
    }

}
