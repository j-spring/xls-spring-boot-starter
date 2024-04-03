package org.jspring.xls.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.jspring.xls.domain.CellSearch;
import org.jspring.xls.utils.CellUtils;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;

/**
 * The XlsxTableService class provides methods for populating an Excel worksheet
 * with data in the form of a table.
 */
public class XlsxTableService {

    /**
     * Updates the provided worksheet with the given data.
     *
     * @param cellSearch  The object used to search for cells in the worksheet.
     * @param cellsPerRow The number of cells per row in the worksheet.
     * @param data        The list of integers containing the data to populate the worksheet with.
     * @return The workbook containing the updated worksheet with the populated data.
     */
    public <T,R> Workbook populateWorksheetWithData(CellSearch<T> cellSearch, int cellsPerRow, List<R> data) {
        processWorksheets(cellsPerRow, data, cellSearch);
        return cellSearch.sheetInfo().workbook();
    }

    private <T,R> void processWorksheets(int cellsPerRow, List<R> data, CellSearch<T> cellSearch) {
        List<Row> rows = new ArrayList<>();
        List<CellData> list = IntStream.range(0, data.size())
                .mapToObj(cellNumber -> createCellData(cellNumber, cellsPerRow, data, cellSearch))
                .toList();

        list.forEach(cellData -> processCellData(cellSearch, rows, cellData));
    }

    private <T,R> CellData createCellData(int index, int cellsPerRow, List<R> data, CellSearch<T> cellSearch) {
        int startRow = cellSearch.cellCoordinates().rowNumber();
        int startColumn = cellSearch.cellCoordinates().columnNumber();

        // Calculate the row and column taking into account the starting row and column
        int rowIndex = index / cellsPerRow + startRow;
        int columnIndex = index % cellsPerRow + startColumn;

        // Make sure column index wraps to zero when it hits `cellsPerRow`
        if (columnIndex >= cellsPerRow + startColumn) {
            columnIndex -= cellsPerRow;
            columnIndex += startColumn;
            rowIndex++;
        }

        CellData<R> cellData = new CellData<>(
                columnIndex,
                rowIndex,
                data.get(index)
        );
        return cellData;
    }

    /*private <T> CellData createCellData(int i, int cellsPerRow, List<Integer> data, CellSearch<T> cellSearch) {
        int startRow = cellSearch.cellCoordinates().rowNumber();
        int startColumn = cellSearch.cellCoordinates().columnNumber();
        int adjustedIndex = i + startColumn + (startRow * cellsPerRow);
        CellData cellData = new CellData(
                adjustedIndex % cellsPerRow,
                adjustedIndex / cellsPerRow,
                data.get(i)
        );
        return cellData;
    }*/

    private <T,R> void processCellData(CellSearch<T> cellSearch, List<Row> rows, CellData<R> cellData) {
        Row row = fetchOrCreateRow(cellSearch, rows, cellData);
        createAndFillCell(row, cellData.columnNumber(), cellData.value());
    }

    private <T,R> Row fetchOrCreateRow(CellSearch<T> cellSearch, List<Row> rows, CellData<R> cellData) {
        //return cellData.columnNumber() == 0 ? addRowToSheet(cellSearch, rows, cellData.rowNumber()) : rows.get(cellData.rowNumber());
        if (cellData.rowNumber() >= rows.size()) {
            return addRowToSheet(cellSearch, rows, cellData.rowNumber());
        }

        return rows.get(cellData.rowNumber());
    }

   /* private <T> Row addRowToSheet(CellSearch<T> cellSearch, List<Row> rows, int rowNumber) {
        Row row = cellSearch.sheetInfo().getSheet().createRow(rowNumber);
        rows.add(row);
        return row;
    }*/

    private <T> Row addRowToSheet(CellSearch<T> cellSearch, List<Row> rows, int rowNumber) {
        if (cellSearch.sheetInfo().getSheet().getRow(rowNumber) == null) {
            Row row = cellSearch.sheetInfo().getSheet().createRow(rowNumber);
            rows.add(row);
            return row;
        } else {
            XSSFRow row = cellSearch.sheetInfo().getSheet().getRow(rowNumber);
            rows.add(row);
            return row;
        }
    }

    private <R> void createAndFillCell(Row row, int columnNumber, R value) {
        Cell cell = row.createCell(columnNumber);
        CellUtils.writeValue(cell, value);
        //cell.setCellValue(value);
    }

    private record CellData<R>(int columnNumber, int rowNumber, R value) {
    }

}