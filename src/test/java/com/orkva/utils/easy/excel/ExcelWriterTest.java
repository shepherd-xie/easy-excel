package com.orkva.utils.easy.excel;

import org.junit.Test;

/**
 * ExcelWriterTest
 *
 * @author Shepherd Xie
 * @version 2022/11/15
 */
public class ExcelWriterTest {

    @Test
    public void testWrite() {
        ExcelWriter.write("./test.xlsx");
    }

}
