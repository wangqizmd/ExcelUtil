package com.ytx.util.util;

import com.ytx.util.annotation.ExcelField;
import com.ytx.util.entity.ExcelParam;
import com.ytx.util.entity.SheetParam;
import com.ytx.util.exception.ExcelException;
import org.apache.commons.collections4.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.*;

/**
 * @author wangqi
 * @version 1.0
 * @className ExcelImportUtil
 * @description TODO
 * @date 2019/7/19 15:32
 */
public class ExcelImportUtil {


    /**
     * 通过文件路径读取excel
     * @param filePath 文件路径
     * @param clazz 接收对象类型
     * @param sheetParams sheet读取参数对象
     * @param <T>  泛型对象
     * @return  读取对象列表
     */
    public static <T> List<T> readExcel(String filePath, Class<T> clazz, SheetParam...sheetParams) {
        return readExcel(new File(filePath),clazz,sheetParams);
    }

    /**
     * 通过文件读取excel
     * @param file 文件
     * @param clazz 接收对象类型
     * @param sheetParams sheet读取参数对象
     * @param <T>  泛型对象
     * @return  读取对象列表
     */
    public static <T> List<T> readExcel(File file, Class<T> clazz,SheetParam...sheetParams) {
        try {
            return readExcel(WorkbookFactory.create(file),clazz,sheetParams);
        } catch (IOException e) {
            throw new ExcelException("导入文件读取失败",e);
        }
    }

    /**
     * 通过字节流读取excel
     * @param bytes 字节流
     * @param clazz 接收对象类型
     * @param sheetParams sheet读取参数对象
     * @param <T>  泛型对象
     * @return  读取对象列表
     */
    public static <T> List<T> readExcel(byte[] bytes, Class<T> clazz, SheetParam...sheetParams) {
        return readExcel(new ByteArrayInputStream(bytes),clazz,sheetParams);
    }

    /**
     * 通过输入流读取excel
     * @param is 输入流
     * @param clazz 接收对象类型
     * @param sheetParams sheet读取参数对象
     * @param <T>  泛型对象
     * @return  读取对象列表
     */
    public static <T> List<T> readExcel(InputStream is, Class<T> clazz, SheetParam...sheetParams) {
        try {
            return readExcel(WorkbookFactory.create(is),clazz,sheetParams);
        } catch (IOException e) {
            throw new ExcelException("导入文件读取失败");
        }finally {
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    throw new ExcelException("文件流关闭失败");
                }
            }
        }
    }

    /**
     * 通过文档对象读取excel
     * @param wb 文档对象
     * @param clazz 接收对象类型
     * @param sheetParams sheet读取参数对象
     * @param <T>  泛型对象
     * @return  读取对象列表
     */
    public static <T> List<T> readExcel(Workbook wb, Class<T> clazz, SheetParam...sheetParams) {
        if(clazz ==null){
            throw new ExcelException("接收对象为空");
        }
        List<T> result = new ArrayList();
        ExcelParam<T> excelParam = new ExcelParam();
        excelParam.setClazz(clazz);
        excelParam.setSheetParams(sheetParams);
        ExcelUtil.getExcelAnnotation(excelParam);
        int numberOfSheets = wb.getNumberOfSheets();
        //根据参数读取sheet，如果sheetIndex为空，读取全部,否则根据数组读取
        if(excelParam.getSheetParams() == null || excelParam.getSheetParams().length == 0){
            for (int sheetNum = 0; sheetNum < numberOfSheets; sheetNum++) {
                //设置当前读取的sheet对象为默认对象
                excelParam.setSheetParam(ExcelUtil.initSheetParam);
                result.addAll(readSheet(wb.getSheetAt(sheetNum), excelParam));
            }
        }else{
            //判断sheet，校验是否为空，是否重复
            Set set = new HashSet();
            for (SheetParam sheetParam: excelParam.getSheetParams()) {
                if(sheetParam == null){
                    continue;
                }
                int index = sheetParam.getSheetIndex();
                String name = sheetParam.getSheetName();
                Sheet sheet = null;
                if(StringUtils.isNotEmpty(name)){
                    sheet = wb.getSheet(name);
                }
                if(sheet==null){
                    if(numberOfSheets<=index){
                        throw new ExcelException("sheetIndex不能越界");
                    }
                    sheet = wb.getSheetAt(index);
                }
                if(!set.add(sheet)){
                    if(StringUtils.isNotEmpty(name)){
                        throw new ExcelException("同时读取多个sheet，sheet不能重复");
                    }else {
                        throw new ExcelException("同时读取多个sheet，sheet Index不能为空或者重复");
                    }
                }
                //设置当前读取的sheet对象
                excelParam.setSheetParam(sheetParam);
                result.addAll(readSheet(sheet,excelParam));
            }
        }
        return result;
    }

    /**
     * 读取sheet
     * @param sheet
     * @param clazz 接收对象类型
     * @param sheetParams sheet参数对象
     * @param <T>  泛型对象
     * @return  读取对象列表
     */
    public static <T> List<T> readSheet(Sheet sheet, Class<T> clazz, SheetParam...sheetParams) {
        ExcelParam<T> excelParam = new ExcelParam();
        excelParam.setClazz(clazz);
        excelParam.setSheetParams(sheetParams);
        excelParam.setSheetParam(ExcelUtil.initSheetParam);
        ExcelUtil.getExcelAnnotation(excelParam);
        if(excelParam.getSheetParams() != null && excelParam.getSheetParams().length != 0){
            excelParam.setSheetParam(excelParam.getSheetParams()[0]);
        }
        return readSheet(sheet, excelParam);
    }

    /**
     * 读取sheet
     * @param sheet
     * @param excelParam  excel参数对象
     * @param <T>  泛型对象
     * @return  读取对象列表
     */
    private static <T> List<T> readSheet(Sheet sheet, ExcelParam<T> excelParam) {
        if(sheet.getLastRowNum()==0){
            return new ArrayList<>();
        }
        SheetParam sheetParam = excelParam.getSheetParam();
        //判断开始读取行数，默认是第一行，如果startIndex存在，则为startIndex，如果startIndex不存在，titleIndex存在，则为titleIndex+1行
        int start = 1;
        if(sheetParam.getStartIndex()!=null && sheetParam.getStartIndex()!= 0){
            start = sheetParam.getStartIndex();
        }else{
            if(sheetParam.getTitleIndex()!=null && sheetParam.getTitleIndex()!= 0){
                start = sheetParam.getTitleIndex() + 1;
            }
        }
        int length = sheet.getLastRowNum();
        //如果存在条数限制，判断条数，以及处理策略
        if(sheetParam.getLength()!=null ){
            if(sheetParam.getLength() < length - (start - 1)){
                if(!sheetParam.isCompatible()){
                    throw new ExcelException("sheet:"+sheet.getSheetName()+"的读取条数超过最大条数限制"+sheetParam.getLength());
                }else{
                    length = sheetParam.getLength() + (start-1);
                }
            }
        }
        List<T> list = new ArrayList<>(length);
        //获取表头
        readTitle(sheet, excelParam);
        // 循环行Row
        for (; start <= length; start++) {
            Row row = sheet.getRow(start);
            if (row == null) {
                continue;
            }
            try{
                T t = readRow(row, excelParam);
                if (t != null) {
                    list.add(t);
                }
            }catch (ExcelException e){
                throw new ExcelException("sheet:"+sheet.getSheetName()+"第" + ( start + 1 )+"行"+e.getMessage());
            }
        }
        return list;
    }

    /**
     * 读取sheet表的标题字段
     * @param sheet
     * @param <T>  泛型对象
     * @return
     */
    private static <T> void readTitle(Sheet sheet, ExcelParam<T> excelParam) {
        SheetParam sheetParam = excelParam.getSheetParam();
        Map<Integer, Field> map = new HashMap<>();
        Map<String, Field> fieldsMap = excelParam.getFieldsMap();
        if (MapUtils.isEmpty(fieldsMap)) {
            throw new ExcelException("excel对象的字段注解为空");
        }
        int titleIndex = 0;
        if(sheetParam.getTitleIndex()!=null && sheetParam.getTitleIndex()!= 0){
            titleIndex = sheetParam.getTitleIndex();
        }
        Row title = sheet.getRow(titleIndex);
        if (title == null) {
            throw new ExcelException("sheet:"+sheet.getSheetName()+"的表头行为空，请检查");
        }
        for (int i = 0; i < title.getLastCellNum(); i++) {
            Cell cell = title.getCell(i);
            if(cell == null){
                throw new ExcelException("sheet:"+sheet.getSheetName()+"的第"+(i+1)+"列表头有误，请检查");
            }
            cell.setCellType(CellType.STRING);
            String val = cell.getStringCellValue();
            if (StringUtils.isEmpty(val.trim())) {
                throw new ExcelException("sheet:"+sheet.getSheetName()+"的第"+(i+1)+"列表头不能为空，请检查");
            }
            if (fieldsMap.get(val) == null) {
                throw new ExcelException("sheet:"+sheet.getSheetName()+"的第"+(i+1)+"列表头'"+title.getCell(i).getStringCellValue()+"'无法匹配，请检查");
            }
            map.put(i, fieldsMap.get(val));
        }
        if(MapUtils.isEmpty(map)){
            throw new ExcelException("sheet:"+sheet.getSheetName()+"无法匹配表头！");
        }
        sheetParam.setTitleMap(map);
    }

    /**
     * 读取excel行
     * @param row
     * @param excelParam  excel参数对象
     * @param <T>  泛型对象
     * @return 读取对象
     */
    private static <T> T readRow(Row row, ExcelParam<T> excelParam) {
        Class<T> clazz = excelParam.getClazz();
        if(clazz ==null){
            throw new ExcelException("接收对象为空");
        }
        SheetParam sheetParam = excelParam.getSheetParam();
        Map<Integer, Field> titleMap = sheetParam.getTitleMap();
        if(MapUtils.isEmpty(titleMap)){
            throw new ExcelException("无法匹配表头！");
        }
        T t = null;
        try {
            t = clazz.newInstance();
        } catch (Exception e) {
            throw new ExcelException("excel指定导入对象为空");
        }
//        if(titleMap.entrySet().size()!=row.getLastCellNum()){
//            throw new ExcelException("列数不匹配");
//        }
        for (int i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            Field field = titleMap.get(i);
            if(field ==null){
                throw new ExcelException("第" + (i + 1) + "列的值无法匹配字段属性");
            }
            //字段是否被忽略
            if(field.getAnnotation(ExcelField.class).ignore()){
                continue;
            }
            //字段是否为空
            if (field.getAnnotation(ExcelField.class).notNull() && (cell == null || ("").equals(cell.toString().trim()))) {
                throw new ExcelException("第" + (i + 1) + "列的属性：" + field.getAnnotation(ExcelField.class).value() + "不能为空");
            }
            if (cell == null) {
                continue;
            }
            Object val = null;
            try{
                val = ExcelUtil.getValue(cell, field);
            }catch (Exception e){
                throw new ExcelException("第" + (i + 1) + "列的属性：" + field.getAnnotation(ExcelField.class).value() + "的值获取失败:"+e.getMessage());
            }
            if (val != null) {
                field.setAccessible(true);
                try {
                    field.set(t, val);
                } catch (Exception e) {
                    throw new ExcelException("第" + (i + 1) + "列的属性：" + field.getAnnotation(ExcelField.class).value() + "的值" + val + "注入失败");
                }
            }
        }
        return t;
    }

}
