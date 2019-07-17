package com.ytx.util;

import com.ytx.util.entity.SheetParam;
import com.ytx.util.entity.TestEntity;
import com.ytx.util.util.ExcelUtil;
import org.junit.Test;

import java.io.File;
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
        List<TestEntity> list = ExcelUtil.readExcel("D:\\Projects\\ExcelUtil\\src\\test\\resources\\test.xlsx", TestEntity.class,sheetParam);
        for (TestEntity entity:list) {
            System.out.println(entity.toString());
        }
    }
}
