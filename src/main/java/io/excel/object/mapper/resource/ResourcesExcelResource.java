package io.excel.object.mapper.resource;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;

/**
 * @author woniper
 */
public class ResourcesExcelResource implements ExcelResource {

    @Override
    public Workbook getResource(String resource) throws IOException {
        return WorkbookFactory.create(findFile(resource));
    }

    private File findFile(String resource) {
        return new File(getClass().getClassLoader().getResource(resource).getFile());
    }
}
