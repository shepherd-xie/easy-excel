package com.orkva.utils.easy.excel;

import com.orkva.utils.easy.excel.annotation.ExcelColumn;
import com.orkva.utils.easy.excel.annotation.ExcelMapper;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.nio.file.Files;
import java.text.MessageFormat;
import java.time.OffsetDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

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
            builder().addSheet(instances).thenBuild().write(Files.newOutputStream(exportFile.toPath()));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    public static ExcelWriterBuilder builder() {
        return builder(ExcelType.XLSX);
    }

    public static ExcelWriterBuilder builder(ExcelType excelType) {
        return new ExcelWriterBuilder(excelType);
    }

    public static final class ExcelSheetBuilder {
        private final ExcelWriterBuilder excelWriterBuilder;
        private final Collection<Object> instances;
        private String title;
        private final Class<?> clazz;

        private ExcelSheetBuilder(Collection<Object> instances, ExcelWriterBuilder excelWriterBuilder) {
            this.instances = instances;
            this.excelWriterBuilder = excelWriterBuilder;
            this.title = "";
            clazz = instances.stream().findAny().get().getClass();
        }

        public ExcelSheetBuilder title(String title) {
            if (title == null || title.isEmpty()) {
                throw new RuntimeException("Sheet title cannot be blank.");
            }
            this.title = title;
            return this;
        }

        public ExcelWriterBuilder then() {
            return excelWriterBuilder;
        }

        public Workbook thenBuild() {
            return then().build();
        }
    }

    public static final class ExcelWriterBuilder {
        /**
         * Excel POI 默认单个字符宽度
         */
        private static final int POI_FONT_WIDTH = 256;
        /**
         * 默认列宽度
         */
        private static final int POI_WIDTH_DEFAULT = 8;
        private static final int POI_SXSS_TRANSFER_COLUMN = 2048;

        private Workbook workbook;
        private final List<ExcelSheetBuilder> excelSheetBuilders;
        private final Map<ExcelColumn, CellStyle> columnCellStyleMap;

        private ExcelWriterBuilder(ExcelType excelType) {
            if (excelType == ExcelType.XLSX) {
                workbook = new XSSFWorkbook();
            } else if (excelType == ExcelType.XLS) {
                workbook = new HSSFWorkbook();
            } else {
                throw new RuntimeException("Unrecognized type.");
            }
            excelSheetBuilders = new ArrayList<>();
            columnCellStyleMap = new HashMap<>();
        }

        public ExcelSheetBuilder addSheet(final Collection instances) {
            if (instances == null || instances.isEmpty()) {
                throw new RuntimeException("Target instances cannot be blank.");
            }
            ExcelSheetBuilder excelSheetBuilder = new ExcelSheetBuilder(instances, this);
            excelSheetBuilders.add(excelSheetBuilder);
            if (instances.size() >= POI_SXSS_TRANSFER_COLUMN && !(workbook instanceof SXSSFSheet)) {
                workbook = new SXSSFWorkbook();
            }
            return excelSheetBuilder;
        }

        public Workbook build() {
            return assembly().workbook;
        }

        private ExcelWriterBuilder assembly() {
            excelSheetBuilders.forEach(this::buildSheet);
            return this;
        }

        private ExcelWriterBuilder buildSheet(final ExcelSheetBuilder excelSheetBuilder) {
            if (!excelSheetBuilder.clazz.isAnnotationPresent(ExcelMapper.class)) {
                throw new RuntimeException("This class is not supported.");
            }

            // 获取有注解的字段
            final List<Field> excelColumnFields = loadExcelColumnField(excelSheetBuilder);
            if (excelColumnFields.isEmpty()) {
                throw new RuntimeException("The target class does not have a valid parse field.");
            }

            final List<ExcelColumn> excelColumns = excelColumnFields.stream()
                    .map(field -> field.getDeclaredAnnotation(ExcelColumn.class))
                    .collect(Collectors.collectingAndThen(Collectors.toList(), Collections::unmodifiableList));

            final Sheet sheet = excelSheetBuilder.title.isEmpty()
                    ? workbook.createSheet()
                    : workbook.createSheet(excelSheetBuilder.title);
            final AtomicInteger index = new AtomicInteger();

            // 设置表头
            final Row header = sheet.createRow(index.getAndIncrement());
            buildHeader(header, excelColumns);

            // 设置内容
            for (Object instance : excelSheetBuilder.instances) {
                if (!excelSheetBuilder.clazz.isInstance(instance)) {
                    throw new RuntimeException("Only allow instances of the same type to be manipulated.");
                }
                buildRow(excelColumnFields, sheet.createRow(index.getAndIncrement()), instance);
            }

            return this;
        }

        private List<Field> loadExcelColumnField(final ExcelSheetBuilder excelSheetBuilder) {
            return Arrays.stream(excelSheetBuilder.clazz.getDeclaredFields())
                    .filter(field -> field.isAnnotationPresent(ExcelColumn.class))
                    .collect(Collectors.collectingAndThen(Collectors.toList(), Collections::unmodifiableList));
        }

        /**
         * Sheet title build
         *
         * @param header
         * @param excelColumns
         * @return
         */
        private Row buildHeader(final Row header, final Collection<ExcelColumn> excelColumns) {
            final AtomicInteger index = new AtomicInteger();
            for (ExcelColumn excelColumn : excelColumns) {
                final Cell cell = header.createCell(index.get(), CellType.STRING);
                cell.setCellValue(excelColumn.value());
                setColumnStyle(header, index.get(), excelColumn);
                index.incrementAndGet();
            }
            return header;
        }

        /**
         * Set column style
         *
         * @param header
         * @param index
         * @param excelColumn
         */
        private void setColumnStyle(final Row header, final int index, final ExcelColumn excelColumn) {
            // TODO: 2022/11/20
        }

        /**
         * Row build
         *
         * @param fields
         * @param row
         * @param instance
         * @return
         */
        private Row buildRow(final Collection<Field> fields, final Row row, final Object instance) {
            final AtomicInteger index = new AtomicInteger();
            for (Field field : fields) {
                final Class<?> typeClass = field.getType();
                final CellType cellType = CellTypeHandler.getCellType(typeClass);
                final Cell cell = buildCell(row.createCell(index.getAndIncrement(), cellType), field, instance);
                cell.setCellStyle(buildCellStyle(field));
            }
            return row;
        }

        /**
         * Cell build
         *
         * @param cell
         * @param field
         * @param instance
         * @return
         */
        private Cell buildCell(final Cell cell, final Field field, final Object instance) {
            final Object value;
            try {
                field.setAccessible(true);
                value = field.get(instance);
            } catch (IllegalAccessException e) {
                throw new RuntimeException("This field is not visible", e);
            } finally {
                field.setAccessible(false);
            }
            if (value == null) {
                return cell;
            }
            WriterCellBuilder<Object> cellBuilder = classCellBuilder(field);
            if (cellBuilder == null) {
                String errorMsg = MessageFormat.format("Cannot find parser for {0}.", field.getType().getSimpleName());
                throw new RuntimeException(errorMsg);
            }
            cellBuilder.build(cell, value);
            return cell;
        }

        /**
         * Get cell builder
         *
         * @param field
         * @return
         */
        private WriterCellBuilder<Object> classCellBuilder(Field field) {
            if (Enum.class.isAssignableFrom(field.getType())) {
                final ExcelColumn excelColumn = field.getDeclaredAnnotation(ExcelColumn.class);
                try {
                    final Field enumValueField = field.getType().getDeclaredField(excelColumn.enumValue());
                    return (cell, value) -> {
                        final Object enumValue;
                        enumValueField.setAccessible(true);
                        try {
                            enumValue = enumValueField.get(value);
                        } catch (IllegalAccessException e) {
                            throw new RuntimeException("This field is not visible", e);
                        } finally {
                            enumValueField.setAccessible(false);
                        }
                        cell.setCellValue((String) enumValue);
                    };
                } catch (NoSuchFieldException e) {
                    return (cell, value) -> {
                        Enum anEnum = (Enum) value;
                        cell.setCellValue(anEnum.name());
                    };
                }
            }
            return CellTypeHandler.getWriterCellBuilder(field.getType());
        }

        /**
         * Cell style build
         *
         * @param field
         * @return
         */
        private CellStyle buildCellStyle(final Field field) {
            final ExcelColumn excelColumn = field.getDeclaredAnnotation(ExcelColumn.class);
            final CellStyle cellStyle = columnCellStyleMap.get(excelColumn);
            if (cellStyle != null) {
                return cellStyle;
            }
            final CellStyle newCellStyle = workbook.createCellStyle();
            if (field.getType() == OffsetDateTime.class) {
                final CreationHelper creationHelper = workbook.getCreationHelper();
                newCellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(DateTimeFormatter.BASIC_ISO_DATE.toString()));
            }
            columnCellStyleMap.put(excelColumn, newCellStyle);
            return newCellStyle;
        }

    }

}
