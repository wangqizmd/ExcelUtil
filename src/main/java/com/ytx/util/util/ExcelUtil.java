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
     * 默认sheet对象
     */
    private static SheetParam initSheetParam;

    /**
     * 初始化默认sheet对象
     */
    static {
        initSheetParam = new SheetParam();
        initSheetParam.setTitleIndex(0);
        initSheetParam.setStartIndex(1);
    }

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
        getExcelAnnotation(excelParam);
        //根据参数读取sheet，如果sheetIndex为空，读取全部,否则根据数组读取
        if(excelParam.getSheetParams() == null || excelParam.getSheetParams().length == 0){
            for (int sheetNum = 0; sheetNum < wb.getNumberOfSheets(); sheetNum++) {
                //设置当前读取的sheet对象为默认对象
                excelParam.setSheetParam(initSheetParam);
                result.addAll(readSheet(wb.getSheetAt(sheetNum), excelParam));
            }
        }else{
            for (int sheetNum = 0; sheetNum < excelParam.getSheetParams().length; sheetNum++){
                //设置当前读取的sheet对象
                excelParam.setSheetParam(excelParam.getSheetParams()[sheetNum]);
                result.addAll(readSheet(wb.getSheetAt(excelParam.getSheetParams()[sheetNum].getSheetIndex()),excelParam));
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
        excelParam.setSheetParam(initSheetParam);
        getExcelAnnotation(excelParam);
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
        if(sheetParam.getLength()!=null && sheetParam.getLength()!=0){
            if(sheetParam.getLength() < sheet.getLastRowNum()){
                throw new ExcelException("sheet:"+sheet.getSheetName()+"的读取条数超过最大条数限制");
            }
        }
        List<T> list = new ArrayList<>(sheet.getLastRowNum());
        //获取表头
        readTitle(sheet, excelParam);
        int start = 1;
        if(sheetParam.getStartIndex()!=null && sheetParam.getStartIndex()!= 0){
            start = sheetParam.getStartIndex();
        }
        // 循环行Row
        for (; start <= sheet.getLastRowNum(); start++) {
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
                throw new ExcelException("sheet:"+sheet.getSheetName()+"的第"+(i+1)+"列表'"+title.getCell(i).getStringCellValue()+"'无法匹配，请检查");
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
            if (field.getAnnotation(ExcelField.class).notNull() && (cell == null || ("").equals(cell.toString().trim()))) {
                throw new ExcelException("第" + (i + 1) + "列的属性：" + field.getAnnotation(ExcelField.class).title() + "不能为空");
            }
            if (cell == null) {
                continue;
            }
            Object val = getValue(cell, field.getType());
            if (val != null) {
                field.setAccessible(true);
                try {
                    field.set(t, val);
                } catch (Exception e) {
                    throw new ExcelException("第" + (i + 1) + "列的属性：" + field.getAnnotation(ExcelField.class).title() + "的值" + val + "注入失败");
                }
            }
        }
        return t;
    }

    /**
     * 读取单元格值
     * @param cell 单元格
     * @param clazz 单元格数据类型
     * @return 单元格值
     */
    private static Object getValue(Cell cell, Class clazz) {
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
    private static<T> void getExcelAnnotation(ExcelParam excelParam) {
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