package com.orkva.utils.easy.excel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.util.List;

/**
 * ExcelReader
 *
 * @author Shepherd Xie
 * @version 2022/11/15
 */
public class ExcelReader {
    private static final Logger logger = LoggerFactory.getLogger(ExcelReader.class);

    /**
     * Read target path excel file to T class list instance.
     *
     * @param path target path for excel file
     * @param clazz entity class of excel record
     * @param <T> entity class
     * @return parsed data list
     */
    public static <T> List<T> read(String path, Class<T> clazz) {
        logger.debug("Read file path: {}", path);
        File sourceFile = new File(path);
        return null;
    }

    public static void main(String[] args) {
        read("", null);
    }

}
