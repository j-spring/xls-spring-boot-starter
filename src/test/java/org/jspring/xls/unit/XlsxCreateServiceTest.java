package org.jspring.xls.unit;

import org.apache.poi.EmptyFileException;
import org.apache.poi.ss.usermodel.*;
import org.jspring.xls.service.XlsxCreateService;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.mockito.Mockito;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import static org.junit.jupiter.api.Assertions.assertThrows;

public class XlsxCreateServiceTest {

    private XlsxCreateService xlsxService;

    @BeforeEach
    public void setUp() {
        xlsxService = new XlsxCreateService();
    }

    @Test
    public void testCreateXlsxFile() throws IOException {
        String fileName = "testFile.xlsx";
        xlsxService.createXlsxFile(fileName);

        try (InputStream inp = new FileInputStream(fileName)) {
            Workbook workbook = WorkbookFactory.create(inp);
            Assertions.assertNotNull(workbook);
        }
    }

    @Test
    public void testWriteToXlsxFile() throws IOException {
        String fileName = "testFile.xlsx";
        String sheetName = "Test Sheet";
        String cellContent = "Test Content";
        int rowIndex = 0;
        int cellIndex = 0;

        xlsxService.createXlsxFile(fileName);
        xlsxService.writeToXlsxFile(fileName, sheetName, rowIndex, cellIndex, cellContent);

        try (InputStream inp = new FileInputStream(fileName)) {
            Workbook workbook = WorkbookFactory.create(inp);
            Sheet sheet = workbook.getSheet(sheetName);
            Assertions.assertNotNull(sheet);

            Row row = sheet.getRow(rowIndex);
            Assertions.assertNotNull(row);

            Cell cell = row.getCell(cellIndex);
            Assertions.assertNotNull(cell);

            String cellValue = cell.getStringCellValue();
            Assertions.assertEquals(cellContent, cellValue);
        }
    }

    @Test
    public void testCreateXlsxFileWithIOException() {
        FileInputStream faultyInputStream = Mockito.mock(FileInputStream.class);
        try {
            Mockito.doThrow(new EmptyFileException()).when(faultyInputStream).read();
        } catch (IOException ignored) {
        }
        assertThrows(EmptyFileException.class,
                () -> xlsxService.createXlsxFileWithInputStream(faultyInputStream));
    }




}
