package io.excel.object.mapper;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;

/**
 * @author woniper
 */
public interface ExcelResource {
    Workbook getResource(String resource) throws IOException;
}
