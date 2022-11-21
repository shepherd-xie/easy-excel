package com.orkva.utils.easy.excel;

import com.orkva.utils.easy.excel.entity.Gender;
import com.orkva.utils.easy.excel.entity.Student;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

/**
 * ExcelWriterTest
 *
 * @author Shepherd Xie
 * @version 2022/11/15
 */
public class ExcelWriterTest {

    @Test
    public void testWrite() {
        List<Student> entities = new ArrayList<>();
        entities.add(new Student(1, "Larry", Gender.MALE));
        entities.add(new Student(2, "Anna", Gender.FEMALE));
        entities.add(new Student(3, "Emma", Gender.FEMALE));
        entities.add(new Student(4, "White", Gender.OTHER));
        ExcelWriter.write(entities, "./src/test/resources/test_export.xlsx");
    }

}
