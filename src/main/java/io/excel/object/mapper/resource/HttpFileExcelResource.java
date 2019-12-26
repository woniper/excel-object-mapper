package io.excel.object.mapper.resource;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.Objects;

/**
 * @author woniper
 */
public class HttpFileExcelResource implements ExcelResource {

    @Override
    public Workbook getResource(String resource) throws IOException {
        URL url = new URL(resource);
        HttpURLConnection connection = (HttpURLConnection)url.openConnection();
        connection.setRequestMethod("GET");

        try (InputStream inputStream = connection.getInputStream()) {
            if (Objects.nonNull(inputStream)) {
                return WorkbookFactory.create(inputStream);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return WorkbookFactory.create(true);
    }

}
