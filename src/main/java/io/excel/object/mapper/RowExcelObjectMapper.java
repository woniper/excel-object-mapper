package io.excel.object.mapper;

import io.excel.object.mapper.resource.ExcelResource;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * @author woniper
 */
public class RowExcelObjectMapper<T extends AbstractRow> implements ExcelObjectMapper<T> {

    private final ExcelResource excelResource;

    public RowExcelObjectMapper(ExcelResource excelResource) {
        this.excelResource = excelResource;
    }

    public List<T> parse(String resource, Class<T> type) {
        return this.parse(resource, 0, 0, type);
    }

    public List<T> parse(String resource, int startRowIndex, Class<T> type) {
        return this.parse(resource, 0, startRowIndex, type);
    }

    public List<T> parse(String resource, int sheetIndex, int startRowIndex, Class<T> type) {
        try (Workbook workbook = excelResource.getResource(resource)) {
            List<T> rows = new ArrayList<>();

            Sheet sheet = workbook.getSheetAt(sheetIndex);

            for (int i = startRowIndex; i < sheet.getPhysicalNumberOfRows(); i++) {
                rows.add(newInstanceRow(type, sheet.getRow(i)));
            }

            return rows;

        } catch (NoSuchMethodException e) {
            // todo Construct(Row)를 찾을 수 없음
            e.printStackTrace();
        } catch (IllegalAccessException | InstantiationException | InvocationTargetException e) {
            // todo newInstance 실패
            e.printStackTrace();
        } catch (IOException e) {
            // todo Workbook 생성 실패
            e.printStackTrace();
        }

        return Collections.emptyList();
    }

    private T newInstanceRow(Class<T> type, Row row) throws NoSuchMethodException, IllegalAccessException, InvocationTargetException, InstantiationException {
        return type.getConstructor(Row.class).newInstance(row);
    }
}
