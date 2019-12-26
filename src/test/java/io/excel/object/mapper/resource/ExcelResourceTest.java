package io.excel.object.mapper.resource;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.IOException;

import static org.assertj.core.api.Assertions.assertThat;

/**
 * @author woniper
 */
public class ExcelResourceTest {

    @Test
    public void excel_파일을_내려받아_Workbook으로_반환() throws IOException {
        // given
        String url = "http://download.microsoft.com/download/1/4/E/14EDED28-6C58-4055-A65C-23B4DA81C4DE/Financial%20Sample.xlsx";
        ExcelResource resource = new HttpFileExcelResource();

        // when
        Workbook workbook = resource.getResource(url);

        // then
        assertThat(workbook.getNumberOfSheets()).isEqualTo(1);
        assertThat(workbook.getSheetAt(0).getPhysicalNumberOfRows()).isEqualTo(701);
    }

    @Test
    public void resources에_있는_excel_파일을_Workbook으로_반환() throws IOException {
        // given
        String fileName = "company.xlsx";
        ExcelResource excelResource = new ResourcesExcelResource();

        // when
        Workbook workbook = excelResource.getResource(fileName);

        // then
        assertThat(workbook.getNumberOfSheets()).isEqualTo(1);
        assertThat(workbook.getSheetAt(0).getPhysicalNumberOfRows()).isEqualTo(4);
    }
}