package org.jspring.xls.unit;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspring.xls.service.XlsxReadingService;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

import static org.junit.jupiter.api.Assertions.*;

class XlsxReadingServiceTest {

    private static final String TEMPLATE_PATH = "src/main/resources/template/Blank.xls";
    private static final String MISSING_TEMPLATE_PATH = "src/main/resources/template/missingTemplate.xlsx";
    private static final String TEMPLATE_BYTE_FILE = "src/main/resources/output/testTemplate.xlsx";
    private static final String SHEET_NAME = "One";

    @Test
    @DisplayName("Test Reading from a Template with Global Path")
    void testReadFromTemplateWithGlobalPath() {
        XlsxReadingService service = new XlsxReadingService(TEMPLATE_PATH);

        XSSFWorkbook book = service.readFromTemplate();

        assertNotNull(book);
        assertEquals(SHEET_NAME, book.getSheetAt(0).getSheetName());
    }

    @Test
    @DisplayName("Test Reading from a Template with Local Path")
    void testReadFromTemplateWithLocalPath() {
        XlsxReadingService service = new XlsxReadingService("");

        XSSFWorkbook book = service.readFromTemplate(TEMPLATE_BYTE_FILE);

        assertNotNull(book);
        assertEquals(SHEET_NAME, book.getSheetAt(0).getSheetName());
    }

    @Test
    @DisplayName("Test Reading from a Template IOException")
    void testReadFromTemplateIOException() {
        XlsxReadingService service = new XlsxReadingService(MISSING_TEMPLATE_PATH);

        Exception exception = assertThrows(RuntimeException.class, service::readFromTemplate);

        assertTrue(exception.getCause() instanceof IOException);
    }

    @Test
    @DisplayName("Test Reading from a Byte Array")
    void testReadFromByteArray() throws IOException {
        byte[] fileContent = Files.readAllBytes(new File(TEMPLATE_BYTE_FILE).toPath());
        XlsxReadingService service = new XlsxReadingService("");

        XSSFWorkbook book = service.readFromByteArray(fileContent);

        assertNotNull(book);
        assertEquals(SHEET_NAME, book.getSheetAt(0).getSheetName());
    }
}