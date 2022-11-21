package com.orkva.utils.easy.excel.annotation;

import java.lang.annotation.*;

/**
 * ExcelColumn
 *
 * @author Shepherd Xie
 * @version 2022/11/17
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelColumn {
    String value();

    String enumValue() default "name";
}
