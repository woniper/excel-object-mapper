package io.excel.object.mapper;

import io.excel.object.mapper.resource.ResourcesExcelResource;
import org.junit.Test;

import java.util.List;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * @author woniper
 */
public class AnnotationExcelObjectMapperTest {

    @Test
    public void employee_excel_row를_AnnotationEmployeeRow로_매핑() {
        // given
        ExcelObjectMapper<AnnotationEmployeeRow> excelObjectMapper = new AnnotationExcelObjectMapper<>(new ResourcesExcelResource());

        // when
        List<AnnotationEmployeeRow> employeeRows = excelObjectMapper.parse("employee.xlsx", 1, AnnotationEmployeeRow.class);

        // then
        assertThat(employeeRows).hasSize(3);
        assertThat(employeeRows.get(0).getName()).isEqualTo("이경원");
        assertThat(employeeRows.get(0).getCompany()).isEqualTo("카페");
        assertThat(employeeRows.get(0).getAge()).isEqualTo(20);
        assertThat(employeeRows.get(0).getShortAge()).isEqualTo((short) 20);
        assertThat(employeeRows.get(0).getDoubleAge()).isEqualTo(20.0);
        assertThat(employeeRows.get(0).getFloatAge()).isEqualTo(20.0f);
        assertThat(employeeRows.get(0).getLongAge()).isEqualTo(20L);
    }

    public static class AnnotationEmployeeRow {

        @CellIndex(index = 0)
        private String name;

        @CellIndex(index = 1)
        private String company;

        @CellIndex(index = 2)
        private int age;

        @CellIndex(index = 2)
        private short shortAge;

        @CellIndex(index = 2)
        private double doubleAge;

        @CellIndex(index = 2)
        private float floatAge;

        @CellIndex(index = 2)
        private long longAge;

        public String getName() {
            return name;
        }

        public String getCompany() {
            return company;
        }

        public int getAge() {
            return age;
        }

        public short getShortAge() {
            return shortAge;
        }

        public double getDoubleAge() {
            return doubleAge;
        }

        public float getFloatAge() {
            return floatAge;
        }

        public long getLongAge() {
            return longAge;
        }
    }
}
