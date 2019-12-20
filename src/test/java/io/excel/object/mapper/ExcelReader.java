package io.excel.object.mapper;

import java.util.List;

/**
 * @author woniper
 */
public interface ExcelReader {
    List<EmployeeRow> read(String resource);
}
