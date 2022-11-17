package com.orkva.utils.easy.excel.entity;

import com.orkva.utils.easy.excel.annotation.ExcelColumn;

/**
 * Student
 *
 * @author Shepherd Xie
 * @version 2022/11/17
 */
public class Student {
    @ExcelColumn("id")
    private Integer id;
    @ExcelColumn("name")
    private String name;
}
