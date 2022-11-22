package com.orkva.utils.easy.excel;

import com.orkva.utils.easy.excel.writer.ExcelWorkbookBuilder;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.util.Collection;

/**
 * ExcelWriter
 *
 * @author Shepherd Xie
 * @version 2022/11/15
 */
public class ExcelWriter {

    /**
     * write excel file to path
     *
     * @param path target path
     */
    public static void write(Collection instances, String path) {
        File exportFile = new File(path);
        try {
            exportFile.createNewFile();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        try {
            builder().addSheet(instances).build().write(Files.newOutputStream(exportFile.toPath()));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    public static ExcelWorkbookBuilder builder() {
        return builder(ExcelType.XLSX);
    }

    public static ExcelWorkbookBuilder builder(ExcelType excelType) {
        return new ExcelWorkbookBuilder(excelType);
    }

}
