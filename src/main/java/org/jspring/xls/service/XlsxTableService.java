package org.jspring.xls.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.jspring.xls.domain.CellSearch;

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
     * @param cellSearch The object used to search for cells in the worksheet.
     * @param cellsPerRow The number of cells per row in the worksheet.
     * @param data The list of integers containing the data to populate the worksheet with.
     * @return The workbook containing the updated worksheet with the populated data.
     */
    public <T> Workbook populateWorksheetWithData(CellSearch<T> cellSearch, int cellsPerRow, List<Integer> data) {
        processWorksheets(cellsPerRow, data, cellSearch);
        return cellSearch.sheetInfo().workbook();
    }

    private <T> void processWorksheets(int cellsPerRow, List<Integer> data, CellSearch<T> cellSearch) {
        List<Row> rows = new ArrayList<>();
        IntStream.range(0, data.size())
                .mapToObj(cellNumber -> createCellData(cellNumber, cellsPerRow, data))
                .forEach(cellData -> processCellData(cellSearch, rows, cellData));
    }

    private CellData createCellData(int i, int cellsPerRow, List<Integer> data) {
        return new CellData(
                i % cellsPerRow,
                i / cellsPerRow,
                data.get(i)
        );
    }

    private <T> void processCellData(CellSearch<T> cellSearch, List<Row> rows, CellData cellData) {
        Row row = fetchOrCreateRow(cellSearch, rows, cellData);
        createAndFillCell(row, cellData.columnNumber(), cellData.value());
    }

    private <T> Row fetchOrCreateRow(CellSearch<T> cellSearch, List<Row> rows, CellData cellData) {
        return cellData.columnNumber() == 0 ? addRowToSheet(cellSearch, rows, cellData.rowNumber()) : rows.get(cellData.rowNumber());
    }

    private <T> Row addRowToSheet(CellSearch<T> cellSearch, List<Row> rows, int rowNumber) {
        Row row = cellSearch.sheetInfo().getSheet().createRow(rowNumber);
        rows.add(row);
        return row;
    }

    private void createAndFillCell(Row row, int columnNumber, int value) {
        Cell cell = row.createCell(columnNumber);
        cell.setCellValue(value);
    }

    private record CellData(int columnNumber, int rowNumber, int value) {
    }

}