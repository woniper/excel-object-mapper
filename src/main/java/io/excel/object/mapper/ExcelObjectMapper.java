package io.excel.object.mapper;

import java.util.List;

/**
 * @author woniper
 */
public interface ExcelObjectMapper<T> {
    List<T> parse(String resource, Class<T> type);

    List<T> parse(String resource, int startRowIndex, Class<T> type);

    List<T> parse(String resource, int sheetIndex, int startRowIndex, Class<T> type);
}
