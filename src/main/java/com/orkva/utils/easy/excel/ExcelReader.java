package com.orkva.utils.easy.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

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
        if (!sourceFile.exists()) {
            throw new RuntimeException("Path file not find.");
        }
        if (!sourceFile.isFile()) {
            throw new RuntimeException("There not a file in the path.");
        }
        String sourceFileName = sourceFile.getName();
        final Workbook workbook;
        if (sourceFileName.endsWith(ExcelConstants.EXCEL_2007_SUFFIX)) {
            try {
                workbook = new XSSFWorkbook(sourceFile);
            } catch (IOException | InvalidFormatException e) {
                throw new RuntimeException(e);
            }
        } else if (sourceFileName.equalsIgnoreCase(ExcelConstants.EXCEL_97_2003_SUFFIX)) {
            try {
                workbook = new HSSFWorkbook(new FileInputStream(sourceFile));
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        } else {
            throw new RuntimeException("Unsupported file.");
        }

        Sheet sheet = workbook.getSheetAt(workbook.getActiveSheetIndex());
        List<T> result = new ArrayList<>();

        for (int rowNum = sheet.getFirstRowNum() + 1; rowNum <= sheet.getLastRowNum(); rowNum ++) {
            final Row row = sheet.getRow(rowNum);
            try {
                Optional.ofNullable(buildInstanceFormRow(clazz.newInstance(), row)).ifPresent(result::add);
            } catch (InstantiationException | IllegalAccessException e) {
                throw new RuntimeException(e);
            }
        }
        return result;
    }

    private static <T> T buildInstanceFormRow(T instance, Row row) {
        if (row == null) {
            throw new RuntimeException("File internal error.");
        }
        if (row.getFirstCellNum() != 0) {
            logger.warn("Find empty line at {}.", row.getRowNum());
            return null;
        }

        for (int cellNum = 0; ; cellNum++) {
            final Cell cell = row.getCell(cellNum, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            final String cellValue = cell.getStringCellValue();
            logger.debug("{}: {}", cell.getAddress(), cellValue);

            if (cellValue == null || cellValue.isEmpty()) {
                break;
            }

        }

        return instance;
    }

    private static void setValue(final Object instance, final Field field, final String value)
            throws IllegalAccessException {
        field.setAccessible(true);

        field.set(instance, value);

        field.setAccessible(false);
    }

}
