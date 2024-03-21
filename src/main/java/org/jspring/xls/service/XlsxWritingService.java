package org.jspring.xls.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;


/**
 * The XlsxWritingService class provides functionality for writing data to an XLSX file.
 */
public class XlsxWritingService {

    /**
     * Writes the contents of a {@link XSSFWorkbook} to a file.
     *
     * @param workbook the {@link XSSFWorkbook} object containing the data to write
     * @param fileName the name of the file to write the data to
     * @throws RuntimeException if an {@link IOException} occurs during the writing process
     */
    public void writeFile(XSSFWorkbook workbook, String fileName) {
        try (FileOutputStream out = new FileOutputStream(fileName)) {
            workbook.write(out);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Writes the contents of a {@link XSSFWorkbook} to a byte array.
     *
     * @param workbook the {@link XSSFWorkbook} object containing the data to write
     * @return a byte array containing the workbook data
     * @throws RuntimeException if an {@link IOException} occurs during the writing process
     */
    public byte[] writeAsByteArray(XSSFWorkbook workbook) {
        try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            workbook.write(baos);
            return baos.toByteArray();

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Writes a value to a cell in an XLSX workbook.
     *
     * @param cell  The cell to write the value to.
     * @param value The value to write to the cell.
     * @throws IllegalStateException If the value type is unexpected.
     */
    public void writeValue(
            Cell cell, Object value) {

        switch (value) {
            case String stringVal -> cell.setCellValue(stringVal);
            case Double doubleVal -> cell.setCellValue(doubleVal);
            case Boolean booleanVal -> cell.setCellValue(booleanVal);
            default -> throw new IllegalStateException("Unexpected value type: " + value.getClass().getName());
        }

    }


}
