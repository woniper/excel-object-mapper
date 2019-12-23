package io.excel.object.mapper;

import io.excel.object.mapper.resource.ExcelResource;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.IOException;
import java.lang.reflect.Field;
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

        try {
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
                        field.setByte(object, numberValue.byteValue());
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
        } catch (IllegalAccessException e) {
            // todo setValue 실패
            e.printStackTrace();
        }
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
