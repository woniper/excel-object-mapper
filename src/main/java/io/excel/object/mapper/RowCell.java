package io.excel.object.mapper;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author woniper
 */
public class RowCell<T> {

    private final Row row;
    private final Cell cell;
    private final int cellNumber;
    private final T value;

    public RowCell(Row row, int cellNumber, Object value) {
        this.row = row;
        this.cell = row.getCell(cellNumber);
        this.cellNumber = cellNumber;
        this.value = (T) value;
    }

    public T getValue() {
        return this.value;
    }
}
