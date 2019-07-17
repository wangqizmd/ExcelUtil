package com.ytx.util.entity;

import com.ytx.util.annotation.Excel;
import com.ytx.util.annotation.ExcelField;
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
public class TestEntity {

    @ExcelField(title = "Id")
    private Integer id;

    @ExcelField(title = "问题名称")
    private String title;

    @ExcelField(title = "一级目录名称")
    private String firstMenu;

    @ExcelField(title = "二级目录名称",notNull = false)
    private String secondMenu;

    @ExcelField(title = "标准答案",ignore = true)
    private String answer;

    @ExcelField(title = "数量")
    private Integer num;

}
