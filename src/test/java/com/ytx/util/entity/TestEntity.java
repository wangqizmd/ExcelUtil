package com.ytx.util.entity;

import com.ytx.util.annotation.Excel;
import com.ytx.util.annotation.ExcelField;
import com.ytx.util.annotation.ExcelFieldChange;
import com.ytx.util.annotation.ExcelSheet;
import lombok.Data;
import lombok.experimental.Accessors;

/**
 * @author wangqi
 * @version 1.0
 * @className KnowledgeCsv
 * @description TODO
 * @date 2019/6/19 11:19
 */
@Data
@Accessors(chain = true)
@Excel(sheet = {
        @ExcelSheet(titleIndex = 2,startIndex = 4,length = 2),
        @ExcelSheet(sheetIndex = 4,sheetName = "sheet2",length = 2,compatible = true)
})
public class TestEntity {

    @ExcelField("Id")
    private Integer id;

    @ExcelField("问题名称")
    private String title;

    @ExcelField("一级目录名称")
    private String firstMenu;

    @ExcelField(value = "二级目录名称",notNull = false,fieldChange = {
            @ExcelFieldChange(key = "false",value = "测试1"),
            @ExcelFieldChange(key = "true",value = "测试2")
    })
    private Boolean secondMenu;

    @ExcelField(value = "标准答案",ignore = true)
    private String answer;

    @ExcelField("数量")
    private Integer num;

}
