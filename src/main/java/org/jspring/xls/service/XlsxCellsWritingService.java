package org.jspring.xls.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class XlsxCellsWritingService {

    private final Workbook workbook;
    private final Sheet sheet;

    public XlsxCellsWritingService(String sheetName) {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet(sheetName);
    }

    public void writeTopToBottom(List<String> values, int maxRows) {
        int colIndex = 0;
        int rowIndex = 0;
        Row row = sheet.createRow(rowIndex);

        for (String value : values) {
            Cell cell = row.createCell(colIndex);
            cell.setCellValue(value);
            rowIndex++;
            if (rowIndex >= maxRows) {
                rowIndex = 0;
                colIndex++;
                row = sheet.createRow(rowIndex);
            }
        }
    }

    public void writeTopToBottom(List<String> values, int maxRows, int maxCols) {
        int rowIndex = 0;
        int colIndex = 0;
        Row row = sheet.createRow(rowIndex);

        for (String value : values) {
            Cell cell = row.createCell(colIndex);
            cell.setCellValue(value);

            rowIndex++;

            if (rowIndex >= maxRows) {
                rowIndex = 0;
                if (++colIndex >= maxCols) {
                    colIndex = 0;
                }
            }

            row = sheet.getRow(rowIndex) != null ? sheet.getRow(rowIndex) : sheet.createRow(rowIndex);
        }
    }

    /*public void writeLeftToRight(List<String> values, int maxCells) {
        int rowIndex = 0;
        int colIndex = 0;
        Row row = sheet.createRow(rowIndex);

        for (String value : values) {
            Cell cell = row.createCell(colIndex);
            cell.setCellValue(value);
            colIndex++;
            if (colIndex >= maxCells) {
                colIndex = 0;
                rowIndex++;
                row = sheet.createRow(rowIndex);
            }
        }
    }*/


}
