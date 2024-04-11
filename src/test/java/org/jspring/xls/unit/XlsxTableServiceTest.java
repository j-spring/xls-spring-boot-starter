package org.jspring.xls.unit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspring.xls.domain.CellCoordinates;
import org.jspring.xls.domain.CellSearch;
import org.jspring.xls.domain.SheetInfo;
import org.jspring.xls.service.XlsxTableService;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class XlsxTableServiceTest {
    private static final String SHEETNAME = "Table";
    private XlsxTableService xlsxTableService;
    private CellSearch<Integer> cellSearch;

    @BeforeEach
    public void setUp() throws IOException {
        xlsxTableService = new XlsxTableService();
        // initialize workbook and create sheet
        XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(true);
        workbook.createSheet(SHEETNAME);
        // create cellSearch object
        cellSearch = new CellSearch<>(
                new SheetInfo(workbook, SHEETNAME),

                new CellCoordinates<>(1, 0, null, null)
        );
    }

    @Test
    @DisplayName("When regular size and data are given, the worksheet should contain a table with ten cells")
    public void populateWorksheetWithData_ShouldReturnTableWithTenCells_GivenRegularSizeAndData() throws Exception {
        List<Integer> data = Arrays.asList(11, 2, 3, 4, 5, 6, 7, 8, 9, 101);
        xlsxTableService.populateWorksheetWithData(cellSearch, 5, data);
        XSSFSheet sheet = cellSearch.sheetInfo().getSheet();
        assertEquals(2, sheet.getPhysicalNumberOfRows(), "Created table should have 10 cells.");

        Row lastRow = sheet.getRow(2);
        assertEquals(5, lastRow.getPhysicalNumberOfCells(), "Last row should have 5 cells.");

        Cell firstCell = lastRow.getCell(0);
        Cell lastCell = lastRow.getCell(4);
        assertEquals(6, firstCell.getNumericCellValue(), "First cell value is wrong");
        assertEquals(101, lastCell.getNumericCellValue(), "Last cell value is wrong");
    }

    private Cell getLastCell() {
        XSSFSheet sheet = cellSearch.sheetInfo().getSheet();
        int lastRowIndex = sheet.getPhysicalNumberOfRows() - 1;
        Row lastRow = sheet.getRow(lastRowIndex);
        int lastCellIndex = lastRow.getPhysicalNumberOfCells() - 1;
        return lastRow.getCell(lastCellIndex);
    }


    @Test
    @DisplayName("When empty data is given, the worksheet should contain an empty table")
    public void populateWorksheetWithData_ShouldReturnEmptyTable_GivenEmptyData() throws Exception {
        List<Integer> data = List.of();
        XSSFSheet sheet = cellSearch.sheetInfo().getSheet();
        xlsxTableService.populateWorksheetWithData(cellSearch, 5, data);
        assertEquals(0, sheet.getPhysicalNumberOfRows(), "Created table should have 0 row.");
    }

    @Test
    @DisplayName("When single cell data is given, the worksheet should contain a table with one cell")
    public void populateWorksheetWithData_ShouldReturnAOneCellTable_GivenSingleCellData() throws Exception {
        List<Integer> data = List.of(1);
        XSSFSheet sheet = cellSearch.sheetInfo().getSheet();
        xlsxTableService.populateWorksheetWithData(cellSearch, 1, data);
        assertEquals(1, sheet.getPhysicalNumberOfRows(), "Created table should have 1 row.");
        Row firstRow = sheet.getRow(1);
        assertEquals(1, firstRow.getPhysicalNumberOfCells(), "First row should have 1 cell.");
        Cell firstCell = firstRow.getCell(0);
        assertEquals(1, firstCell.getNumericCellValue(), "Cell value is wrong");
    }

    @Test
    @DisplayName("When single cell data is given and a cellSearch is set, the worksheet should contain a table with one cell starting from the cellSearch")
    public void populateWorksheetWithData_ShouldReturnAOneCellTableAtRightPosition_GivenSingleCellData() throws Exception {
        List<Integer> data = List.of(12);
        // customize cellSearch
        SheetInfo sheetInfo = cellSearch.sheetInfo();
        cellSearch = new CellSearch<>(
                sheetInfo,
                new CellCoordinates<>(1, 3, null, null)
        );

        XSSFSheet sheet = cellSearch.sheetInfo().getSheet();
        xlsxTableService.populateWorksheetWithData(cellSearch, 1, data);
        assertEquals(1, sheet.getPhysicalNumberOfRows(), "Created table should have 1 row.");
        Row firstRow = sheet.getRow(1);
        assertEquals(1, firstRow.getPhysicalNumberOfCells(), "First row should have 1 cell.");
        Cell firstCell = firstRow.getCell(3);
        assertEquals(12, firstCell.getNumericCellValue(), "Cell value is wrong");
    }

    @Test
    @DisplayName("When multiple cell data are given and a cellSearch is set, the worksheet should contain a table with multiple cell starting from the cellSearch")
    public void populateWorksheetWithData_ShouldReturnAOneCellTableAtRightPosition_GivenMultipleCellData() throws Exception {
        List<Integer> data = List.of(12, 1, 56, 78, 23, 3, 67, 9);
        // customize cellSearch
        SheetInfo sheetInfo = cellSearch.sheetInfo();
        cellSearch = new CellSearch<>(
                sheetInfo,
                new CellCoordinates<>(2, 3, null, null)
        );

        XSSFSheet sheet = cellSearch.sheetInfo().getSheet();
        xlsxTableService.populateWorksheetWithData(cellSearch, 3, data);
        assertEquals(3, sheet.getPhysicalNumberOfRows(), "Created table should have 3 rows.");
        Row lastRow = sheet.getRow(4);
        assertEquals(2, lastRow.getPhysicalNumberOfCells(), "Last row should have 2 cells.");
        Cell lastCell = lastRow.getCell(4);
        assertEquals(9, lastCell.getNumericCellValue(), "Cell value is wrong");
    }

    @Test
    @DisplayName("Write table with multiple cell with double data")
    public void populateWorksheetWithMultipleDoubleData() throws Exception {
        List<Double> data = List.of(12.3, 1.5, 56.6, 78.0, 23.0, 3.1, 67.2, 9.5);
        // customize cellSearch
        SheetInfo sheetInfo = cellSearch.sheetInfo();
        cellSearch = new CellSearch<>(
                sheetInfo,
                new CellCoordinates<>(2, 3, null, null)
        );

        XSSFSheet sheet = cellSearch.sheetInfo().getSheet();
        xlsxTableService.populateWorksheetWithData(cellSearch, 3, data);
        assertEquals(3, sheet.getPhysicalNumberOfRows(), "Created table should have 3 rows.");
        Row lastRow = sheet.getRow(4);
        assertEquals(2, lastRow.getPhysicalNumberOfCells(), "Last row should have 2 cells.");
        Cell lastCell = lastRow.getCell(4);
        assertEquals(9.5, lastCell.getNumericCellValue(), "Cell value is wrong");
    }

    @Test
    @DisplayName("Write table with multiple cell with String data")
    public void populateWorksheetWithMultipleStringData() throws Exception {
        List<String> data = List.of("12.3", "1.5", "56.6", "78.0", "23.0", "3.1", "67.2", "9.4");
        // customize cellSearch
        SheetInfo sheetInfo = cellSearch.sheetInfo();
        cellSearch = new CellSearch<>(
                sheetInfo,
                new CellCoordinates<>(2, 3, null, null)
        );

        XSSFSheet sheet = cellSearch.sheetInfo().getSheet();
        xlsxTableService.populateWorksheetWithData(cellSearch, 3, data);
        assertEquals(3, sheet.getPhysicalNumberOfRows(), "Created table should have 3 rows.");
        Row lastRow = sheet.getRow(4);
        assertEquals(2, lastRow.getPhysicalNumberOfCells(), "Last row should have 2 cells.");
        Cell lastCell = lastRow.getCell(4);
        assertEquals("9.4", lastCell.getStringCellValue(), "Cell value is wrong");
    }

    @Test
    @DisplayName("Write table with multiple cell with double data and a two-cell depth")
    public void populateWorksheetWithMultipleDoubleDataAndDepth() throws Exception {
        List<Double> data = List.of(12.3, 1.5, 56.6, 78.0, 23.0, 3.1, 67.2, 9.5);
        // customize cellSearch
        SheetInfo sheetInfo = cellSearch.sheetInfo();
        cellSearch = new CellSearch<>(
                sheetInfo,
                new CellCoordinates<>(2, 3, null, null)
        );

        XSSFSheet sheet = cellSearch.sheetInfo().getSheet();
        xlsxTableService.populateWorksheetWithData(cellSearch, 3, data);
        assertEquals(3, sheet.getPhysicalNumberOfRows(), "Created table should have 3 rows.");
        Row lastRow = sheet.getRow(4);
        assertEquals(2, lastRow.getPhysicalNumberOfCells(), "Last row should have 2 cells.");
        Cell lastCell = lastRow.getCell(4);
        assertEquals(9.5, lastCell.getNumericCellValue(), "Cell value is wrong");
    }

}