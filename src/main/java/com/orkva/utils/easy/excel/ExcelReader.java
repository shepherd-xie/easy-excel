package com.orkva.utils.easy.excel;

import com.orkva.utils.easy.excel.annotation.ExcelColumn;
import com.orkva.utils.easy.excel.parser.ExcelClassParser;
import com.orkva.utils.easy.excel.parser.ExcelClassParserFactory;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.text.MessageFormat;
import java.util.*;
import java.util.stream.Collectors;

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

        Field[] declaredFields = instance.getClass().getDeclaredFields();
        List<Field> parserFields = Arrays.stream(declaredFields)
                .filter(field -> field.isAnnotationPresent(ExcelColumn.class))
                .collect(Collectors.toList());

        for (int cellNum = 0; cellNum < parserFields.size(); cellNum ++) {
            Field field = parserFields.get(cellNum);
            Cell cell = row.getCell(cellNum, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            final String cellValue = getCellValueAsString(cell);
            logger.debug("{}=>{}: {}", field.getName(), cell.getAddress(), cellValue);

            if (cellValue == null || cellValue.isEmpty()) {
                continue;
            }

            try {
                setValue(instance, field, cellValue);
            } catch (IllegalAccessException e) {
                throw new RuntimeException(e);
            }
        }

        return instance;
    }

    private static void setValue(Object instance, Field field, String value)
            throws IllegalAccessException {
        if (value == null || value.isEmpty()) {
            return;
        }
        field.setAccessible(true);

        if (field.getType().isEnum()) {
            Object eEnum = enumParser(field.getType(), value);
            field.set(instance, eEnum);
        } else {
            ExcelClassParser<?> parser = ExcelClassParserFactory.getParser(field.getType());
            if (parser == null) {
                field.setAccessible(false);
                String errorMsg = MessageFormat.format("Cannot find parser for {0} ", field.getType().getSimpleName());
                logger.error(errorMsg);
                throw new RuntimeException(errorMsg);
            }
            field.set(instance, parser.parse(value));
        }

        field.setAccessible(false);
    }

    private static final DecimalFormat NUMERIC_FORMATTER = new DecimalFormat("#.##");

    private static String getCellValueAsString(Cell cell) {
        final String cellValue;
        if (CellType.NUMERIC == cell.getCellType()) {
            cellValue = NUMERIC_FORMATTER.format(cell.getNumericCellValue());
        } else {
            cellValue = cell.getStringCellValue();
        }
        return cellValue;
    }

    private static Object enumParser(Class<?> enumClass, String value) {
        final Object[] enumConstants = enumClass.getEnumConstants();
        final Field nameField;
        try {
            nameField = enumClass.getDeclaredField("name");
        } catch (NoSuchFieldException e) {
            throw new RuntimeException("Cannot find attribute 'name' in Enum.", e);
        }
        nameField.setAccessible(true);
        for (Object enumConstant : enumConstants) {
            final Object name;
            try {
                name = nameField.get(enumConstant);
            } catch (IllegalAccessException e) {
                throw new RuntimeException(e);
            }
            if (Objects.equals(name, value)) {
                return enumConstant;
            }
        }
        String enumNotFound = MessageFormat
                .format("enumClass[{0}], value[{1}] cannot find pair enum", enumClass, value);
        throw new RuntimeException(enumNotFound);
    }

}
