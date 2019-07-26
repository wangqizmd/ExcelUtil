package com.ytx.util;

import com.ytx.util.entity.SheetParam;
import com.ytx.util.entity.TestEntity;
import com.ytx.util.util.ExcelImportUtil;
import org.junit.Test;

import java.util.List;

/**
 * @author wangqi
 * @version 1.0
 * @className ExcelImportTest
 * @description TODO
 * @date 2019/7/17 17:01
 */
public class ExcelImportTest {

    @Test
    public void downLoadTest() {
        SheetParam sheetParam = new SheetParam();
        sheetParam.setSheetIndex(0);
        sheetParam.setTitleIndex(2);
        sheetParam.setStartIndex(3);
        sheetParam.setLength(2);
        sheetParam.setCompatible(true);
        SheetParam sheetParam1 = new SheetParam();
//        sheetParam1.setSheetIndex(4);
        sheetParam1.setSheetName("sheet2");
//        sheetParam1.setTitleIndex(0);
        sheetParam1.setStartIndex(1);
        sheetParam1.setLength(3);
        List<TestEntity> list = ExcelImportUtil.readExcel("D:\\Projects\\ExcelUtil\\src\\test\\resources\\test.xlsx",
                TestEntity.class,sheetParam,sheetParam1,null);
        for (TestEntity entity:list) {
            System.out.println(entity.toString());
        }
    }
}
