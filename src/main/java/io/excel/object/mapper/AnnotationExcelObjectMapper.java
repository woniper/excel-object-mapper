package io.excel.object.mapper;

import io.excel.object.mapper.resource.ExcelResource;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.*;
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
            Sheet sheet = excelResource.getResource(resource).getSheetAt(sheetIndex);

            for (int i = startRowIndex; i < sheet.getPhysicalNumberOfRows(); i++) {
                T object = type.newInstance();
                Row row = sheet.getRow(i);

                Arrays.stream(object.getClass().getDeclaredFields())
                        .peek(x -> x.setAccessible(true))
                        .forEach(x -> setFieldValue(row, object, x));

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

    private void setFieldValue(Row row, T object, Field field) {
        if (notAvailableAccessField(field)) {
            return;
        }

        CellIndex cellIndex = field.getDeclaredAnnotation(CellIndex.class);
        Cell cell = row.getCell(cellIndex.index());

        if (Objects.isNull(cell)) {
            throw new IllegalArgumentException(cellIndex.message());
        }

        cell.setCellType(CellType.STRING);

        if (field.getType().equals(RowCell.class)) {
            Object value = getPrimitiveTypeValue(cell.getStringCellValue(), getRowCellGenericType(field));
            RowCell<T> rowCell = new RowCell<>(row, cellIndex.index(), value);
            setValue(field, object, rowCell);
        } else {
            Object value = getPrimitiveTypeValue(cell.getStringCellValue(), field.getGenericType());
            setValue(field, object, value);
        }
    }

    private Type getRowCellGenericType(Field field) {
        return ((ParameterizedType) field.getGenericType()).getActualTypeArguments()[0];
    }

    private void setValue(Field field, T object, Object value) {
        try {
            field.set(object, value);
        } catch (IllegalAccessException e) {
            // todo setValue 실패
            e.printStackTrace();
        }
    }

    private Object getPrimitiveTypeValue(String value, Type fieldType) {
        String typeName = fieldType.getTypeName();

        if (isNumeric(value)) {
            Number numberValue = Double.valueOf(value);

            switch (typeName) {
                case "int":
                case "java.lang.Integer":
                    return numberValue.intValue();
                case "byte":
                case "java.lang.Byte":
                    return numberValue.byteValue();
                case "short":
                case "java.lang.Short":
                    return numberValue.shortValue();
                case "double":
                case "java.lang.Double":
                    return numberValue.doubleValue();
                case "float":
                case "java.lang.Float":
                    return numberValue.floatValue();
                case "long":
                case "java.lang.Long":
                    return numberValue.longValue();
            }
        } else if (typeName.equals("java.lang.Boolean") || typeName.equals("boolean")) {
            return Boolean.parseBoolean(value);
        }

        return value;
    }

    private boolean notAvailableAccessField(Field field) {
        return !field.isAnnotationPresent(CellIndex.class);
    }

    private boolean isNumeric(String value) {
        if (Objects.isNull(value)) {
            return false;
        }

        return pattern.matcher(value).matches();
    }
}
