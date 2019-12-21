package io.excel.object.mapper;

import org.apache.poi.ss.usermodel.Row;

/**
 * @author woniper
 */
public abstract class AbstractRow {
    private final Row row;

    protected AbstractRow(Row row) {
        this.row = row;
    }
}
