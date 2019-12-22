package io.excel.object.mapper;

import io.excel.object.mapper.resource.ResourcesExcelResource;
import org.junit.Test;

import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * @author woniper
 */
public class RowExcelObjectMapperTest {

    @Test
    public void Row_파라미터를_받는_생성자가_없는_object() {
        // given
        class TestRow extends AbstractRow {
            protected TestRow() {
                super(null);
            }
        }

        ExcelObjectMapper<TestRow> parser = new RowExcelObjectMapper<>(new ResourcesExcelResource());

        // when
        List<TestRow> testRows = parser.parse("employee.xlsx", 1, TestRow.class);

        // then
        assertThat(testRows).hasSize(0);
    }

    @Test
    public void employee_excel_row를_EmployeeRow로_매핑() {
        // given
        ExcelObjectMapper<EmployeeRow> parser = new RowExcelObjectMapper<>(new ResourcesExcelResource());

        // when
        List<EmployeeRow> employeeRows = parser.parse("employee.xlsx", 1, EmployeeRow.class);

        // then
        assertThat(employeeRows).hasSize(3);
        assertThat(employeeRows.get(0).getName()).isEqualTo("이경원");
        assertThat(employeeRows.get(0).getCompany()).isEqualTo("카페");
        assertThat(employeeRows.get(0).getAge()).isEqualTo(20);
    }

    @Test
    public void company_excel_row를_CompanyRow로_매핑() {
        // given
        ExcelObjectMapper<CompanyRow> parser = new RowExcelObjectMapper<>(new ResourcesExcelResource());

        // when
        List<CompanyRow> companies = parser.parse("company.xlsx", 1, CompanyRow.class);

        // then
        assertThat(companies).hasSize(3);
        assertThat(companies.get(0).getCompany()).isEqualTo("카페");
        assertThat(companies.get(0).getAddress()).isEqualTo("서울");
        assertThat(companies.get(0).getEmployeeCount()).isEqualTo(10);
    }
}
