package io.excel.object.mapper;

import io.excel.object.mapper.resource.ExcelResource;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;
import java.util.regex.Pattern;

/**
 * @author woniper
 */
public class AnnotationExcelObjectMapper<T> implements ExcelObjectMapper<T> {

    private final ExcelResource excelResource;
    private Pattern pattern = Pattern.compile("-?\\d+(\\.\\d+)?");

    public AnnotationExcelObjectMapper(ExcelResource excelResource) {
        this.excelResource = excelResource;
    }

    @Override
    public List<T> parse(String resource, Class<T> type) {
        return this.parse(resource, 0, 0, type);
    }

    @Override
    public List<T> parse(String resource, int startRowIndex, Class<T> type) {
        return this.parse(resource, 0, startRowIndex, type);
    }

    @Override
    public List<T> parse(String resource, int sheetIndex, int startRowIndex, Class<T> type) {
        try {
            List<T> rows = new ArrayList<>();
            Sheet sheet = getSheet(resource, sheetIndex);

            for (int i = startRowIndex; i < sheet.getPhysicalNumberOfRows(); i++) {
                T object = type.newInstance();
                for (Field field : object.getClass().getDeclaredFields()) {
                    if (availableAccessField(field)) {
                        CellIndex cellIndex = field.getDeclaredAnnotation(CellIndex.class);

                        Cell cell = getCell(sheet.getRow(i), cellIndex.index());

                        if (Objects.isNull(cell)) {
                            throw new IllegalArgumentException(cellIndex.message());
                        }

                        setFieldValue(object, field, cell);
                    }
                }

                rows.add(object);
            }

            return rows;

        } catch (IOException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InstantiationException e) {
            e.printStackTrace();
        }

        return Collections.emptyList();
    }

    private void setFieldValue(T object, Field field, Cell cell) throws IllegalAccessException {
        field.setAccessible(true);
        cell.setCellType(CellType.STRING);
        String cellValue = cell.getStringCellValue();
        String typeName = field.getType().getCanonicalName();

        if (isNumeric(cellValue)) {
            Number numberValue = Double.valueOf(cellValue);
            switch (typeName) {
                case "int":
                    field.setInt(object, numberValue.intValue());
                    break;
                case "byte":
                    field.setInt(object, numberValue.byteValue());
                    break;
                case "short":
                    field.setShort(object, numberValue.shortValue());
                    break;
                case "double":
                    field.setDouble(object, numberValue.doubleValue());
                    break;
                case "float":
                    field.setFloat(object, numberValue.floatValue());
                    break;
                case "long":
                    field.setLong(object, numberValue.longValue());
                    break;
            }

        } else {
            if (typeName.equals("boolean")) {
                field.setBoolean(object, Boolean.parseBoolean(cellValue));
            } else {
                field.set(object, cellValue);
            }
        }
    }

    private Cell getCell(Row row, int index) {
        return row.getCell(index);
    }

    private boolean availableAccessField(Field field) {
        return field.isAnnotationPresent(CellIndex.class);
    }

    private Sheet getSheet(String resource, int sheetIndex) throws IOException {
        return excelResource.getResource(resource).getSheetAt(sheetIndex);
    }

    private boolean isNumeric(String value) {
        if (Objects.isNull(value)) {
            return false;
        }

        return pattern.matcher(value).matches();
    }
}
