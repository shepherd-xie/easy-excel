package com.orkva.utils.easy.excel.entity;

import com.orkva.utils.easy.excel.annotation.ExcelColumn;
import com.orkva.utils.easy.excel.annotation.ExcelMapper;

/**
 * Student
 *
 * @author Shepherd Xie
 * @version 2022/11/17
 */
@ExcelMapper("Student")
public class Student {
    @ExcelColumn("id")
    private Integer id;
    @ExcelColumn("name")
    private String name;

    public Student(Integer id, String name) {
        this.id = id;
        this.name = name;
    }
}
