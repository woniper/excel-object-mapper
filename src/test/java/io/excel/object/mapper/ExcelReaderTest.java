package io.excel.object.mapper;

import org.junit.Test;

import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * @author woniper
 */
public class ExcelReaderTest {

    @Test
    public void user_excel_row를_user로_파싱() {
        ResourceExcelParser parser = new ResourceExcelParser();
        List<EmployeeRow> employeeRows = parser.parse("user.xlsx", EmployeeRow.class);

        assertThat(employeeRows).hasSize(3);
        assertThat(employeeRows.get(0).getName()).isEqualTo("이경원");
        assertThat(employeeRows.get(0).getCompany()).isEqualTo("카페");
        assertThat(employeeRows.get(0).getAge()).isEqualTo(20);
    }

    @Test
    public void company_excel_row를_company로_파싱() {
        ResourceExcelParser parser = new ResourceExcelParser();
        List<CompanyRow> companies = parser.parse("company.xlsx", CompanyRow.class);

        assertThat(companies).hasSize(3);
        assertThat(companies.get(0).getCompany()).isEqualTo("카페");
        assertThat(companies.get(0).getAddress()).isEqualTo("서울");
        assertThat(companies.get(0).getEmployeeCount()).isEqualTo(10);
    }
}
