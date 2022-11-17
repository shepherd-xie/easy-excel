package com.orkva.utils.easy.excel;

import com.orkva.utils.easy.excel.entity.Student;
import org.junit.Assert;
import org.junit.Test;

/**
 * ExcelWriterTest
 *
 * @author Shepherd Xie
 * @version 2022/11/15
 */
public class ExcelReaderTest {

    @Test
    public void testRead() {
        Assert.assertNotNull(ExcelReader.read(ClassLoader.getSystemResource("test_student.xlsx").getFile(), Student.class));
    }

}
