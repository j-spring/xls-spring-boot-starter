package org.jspring.xls.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.jspring.xls.domain.SheetInfo;
import org.jspring.xls.domain.StartPoint;
import org.jspring.xls.domain.TableData;
import org.jspring.xls.utils.CellUtils;

import java.util.List;

public class XlsxCellsWritingService {

   /* private final Workbook workbook;
    private final Sheet sheet;

    public XlsxCellsWritingService(String sheetName) {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet(sheetName);
    }*/


    public <T> void writeTopToBottom(
            SheetInfo sheetInfo,
            StartPoint startPoint,
            TableData<T> tableData
    ) {

        Sheet sheet = sheetInfo.getSheet();
        List<T> values = tableData.values();
        int maxRows = tableData.maxRows();
        int maxCols = tableData.maxCols();

        // Calculate the maximum size and resize the list of values if necessary
        int maxSize = maxRows * maxCols;
        List<T> resizedValues = values.size() > maxSize ? values.subList(0, maxSize) : values;

        int currentRow = startPoint.startRow();
        int currentCol = startPoint.startColumn();
        for (T value : resizedValues) {

            Row row = sheet.getRow(currentRow);
            if (row == null) {
                row = sheet.createRow(currentRow);
            }
            Cell newCell = row.createCell(currentCol);
            CellUtils.writeValue(newCell, value);

            currentRow++;
            if (currentRow >= startPoint.startRow() + maxRows) {
                currentRow = startPoint.startRow();
                currentCol++;
                if (currentCol >= startPoint.startColumn() + maxCols) {
                    currentCol = startPoint.startColumn(); // Wrap around to the initial column
                }
            }
        }
    }


   /* public void writeTopToBottom(List<String> values, int maxRows) {
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
    }*/

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
