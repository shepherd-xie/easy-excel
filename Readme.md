# Easy Excel

基于 `Apache POI` 的注解驱动的 Excel 解析工具。

### 导出数据到 Excel：

```java
Collection<Object> instances = new List<>();
File exportFile = new File("/path/to/excel");
exportFile.createNewFile();
ExcelWriter.write(instances, Files.newOutputStream(exportFile.toPath()));
```

### 从 Excel 读取数据到程序中：

1. 定义 Excel 解析后的实体类
```java
@ExcelMapper("Student")
public class Student {
    @ExcelColumn("id")
    private Integer id;
    @ExcelColumn("name")
    private String name;
    @ExcelColumn("gander")
    private Gender gender;
    @ExcelColumn("birth")
    private OffsetDateTime birth;
}
```
2. 解析 Excel 返回实体类对象
```java
File file = new File(ClassLoader.getSystemResource("/path/to/excel").getFile());
List<Student> students = ExcelReader.read(file, Student.class);
```

### 自定义类型解析

1. 实现 `com.orkva.utils.easy.excel.parser.ExcelClassParser` 接口
```java
public class OffsetDateTimeExcelClassParser implements ExcelClassParser<OffsetDateTime> {

    @Override
    public OffsetDateTime parse(String value) {
        return LocalDate.parse(value, DateTimeFormatter.ofPattern("y/M/d").withZone(ZoneId.of("Asia/Shanghai")))
                .atTime(LocalTime.MIDNIGHT)
                .atZone(ZoneId.of("Asia/Shanghai")).toOffsetDateTime();
    }

}
```
2. 在 `ExcelClassParserRegister` 中注册
```java
ExcelClassParserRegister.register(OffsetDateTime.class, new OffsetDateTimeExcelClassParser());
```