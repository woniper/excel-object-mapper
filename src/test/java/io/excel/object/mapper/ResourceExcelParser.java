package io.excel.object.mapper;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 * @author woniper
 */
public class ResourceExcelParser {

    public <T extends AbstractRow> List<T> parse(String resource, Class<T> type) {
        try {
            List<T> rows = new ArrayList<>();
            Sheet sheet = getSheet(resource);

            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                rows.add(newInstanceRow(type, sheet.getRow(i)));
            }

            return rows;

        } catch (IOException e) {
            e.printStackTrace();
        } catch (NoSuchMethodException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }

        return Collections.emptyList();
    }

    private <T> T newInstanceRow(Class<T> type, Row row) throws InstantiationException, IllegalAccessException, InvocationTargetException, NoSuchMethodException {
        return type.getConstructor(Row.class).newInstance(row);
    }

    private Sheet getSheet(String resource) throws IOException, InvalidFormatException {
        Workbook workbook = new XSSFWorkbook(getFile(resource));
        return workbook.getSheetAt(0);
    }

    private File getFile(String resource) {
        return new File(getClass().getClassLoader().getResource(resource).getFile());
    }

}
