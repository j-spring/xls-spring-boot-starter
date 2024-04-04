package org.jspring.xls.service;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * A service class for creating and writing to XLSX files.
 */
public class XlsxCreateService {

    /**
     * Creates a new XLSX file with the specified fileName.
     *
     * @param fileName The name of the new XLSX file.
     * @throws IOException If an I/O error occurs.
     */
    public void createXlsxFile(String fileName) throws IOException {
        Workbook workbook = createWorkbook();
        writeFile(workbook, fileName);
    }

    /**
     * Writes content to a specific cell in an XLSX file.
     *
     * @param fileName     The name of the XLSX file to write to.
     * @param sheetName    The name of the sheet to write to.
     * @param rowIndex     The index of the row to write to.
     * @param cellIndex    The index of the cell to write to.
     * @param cellContent  The content to write to the cell.
     * @throws IOException If an I/O error occurs.
     */
    public void writeToXlsxFile(String fileName, String sheetName, int rowIndex, int cellIndex, String cellContent) throws IOException {
        try (FileInputStream fis = new FileInputStream(fileName)) {
            Workbook workbook = WorkbookFactory.create(fis);
            Sheet sheet = getOrCreateSheet(workbook, sheetName);
            Row row = getOrCreateRow(sheet, rowIndex);
            Cell cell = getOrCreateCell(row, cellIndex);
            cell.setCellValue(cellContent);
            writeFile(workbook, fileName);
        }
    }

    /**
     * Creates an XLSX file with the specified InputStream as the source.
     *
     * @param faultyInputStream The InputStream to read from. It should contain valid XLSX data.
     * @throws IOException If an I/O error occurs while processing the InputStream.
     */
    public void createXlsxFileWithInputStream(InputStream faultyInputStream) throws IOException {
        Workbook workbook = WorkbookFactory.create(faultyInputStream);
        writeFile(workbook, "faultyFile.xlsx");
    }

    private Workbook createWorkbook() throws IOException {
        return WorkbookFactory.create(true);
    }

    private void writeFile(Workbook workbook, String fileName) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(fileName)) {
            workbook.write(fos);
        }
    }

    private Sheet getOrCreateSheet(Workbook workbook, String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            sheet = workbook.createSheet(sheetName);
        }
        return sheet;
    }

    private Row getOrCreateRow(Sheet sheet, int index) {
        Row row = sheet.getRow(index);
        if (row == null) {
            row = sheet.createRow(index);
        }
        return row;
    }

    // isNewly added method
    private Cell getOrCreateCell(Row row, int index) {
        Cell cell = row.getCell(index);
        if (cell == null) {
            cell = row.createCell(index);
        }
        return cell;
    }
}

