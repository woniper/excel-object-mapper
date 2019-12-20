package io.excel.object.mapper;

import org.apache.poi.ss.usermodel.Row;

/**
 * @author woniper
 */
public class CompanyRow extends AbstractRow {
    private String company;
    private String address;
    private int employeeCount;

    public CompanyRow(Row row) {
        super(row);
        this.company = row.getCell(0).getStringCellValue();
        this.address = row.getCell(1).getStringCellValue();
        this.employeeCount = (int)row.getCell(2).getNumericCellValue();
    }

    public String getCompany() {
        return company;
    }

    public String getAddress() {
        return address;
    }

    public int getEmployeeCount() {
        return employeeCount;
    }
}
