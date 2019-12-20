package io.excel.object.mapper;

import org.apache.poi.ss.usermodel.Row;

/**
 * @author woniper
 */
public class EmployeeRow extends AbstractRow {

    private String name;
    private String company;
    private int age;

    public EmployeeRow(Row row) {
        super(row);
        this.name = row.getCell(0).getStringCellValue();
        this.company = row.getCell(1).getStringCellValue();
        this.age = (int)row.getCell(2).getNumericCellValue();
    }

    public String getName() {
        return name;
    }

    public String getCompany() {
        return company;
    }

    public int getAge() {
        return age;
    }
}
