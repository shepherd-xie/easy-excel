package com.orkva.utils.easy.excel;

import com.orkva.utils.easy.excel.entity.Student;
import org.junit.Assert;
import org.junit.Test;

import java.util.List;

/**
 * ExcelWriterTest
 *
 * @author Shepherd Xie
 * @version 2022/11/15
 */
public class ExcelReaderTest {

    @Test
    public void testRead() {
        List<Student> students = ExcelReader.read(ClassLoader.getSystemResource("test_student.xlsx").getFile(), Student.class);
        for (Student student : students) {
            System.out.println(student);
        }
        Assert.assertNotNull(students);
    }

}
