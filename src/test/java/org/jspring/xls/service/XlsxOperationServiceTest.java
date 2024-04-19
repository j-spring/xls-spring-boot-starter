package org.jspring.xls.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspring.xls.builder.WorkbookOperation;
import org.jspring.xls.domain.CellCoordinates;
import org.jspring.xls.domain.SheetInfo;
import org.jspring.xls.domain.TableData;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.*;
import java.util.AbstractMap.SimpleEntry;

import static org.jspring.xls.builder.WorkbookOperation.*;
import static org.jspring.xls.domain.CellCoordinates.*;
import static org.junit.jupiter.api.Assertions.*;

public class XlsxOperationServiceTest {

    private static final String SHEET_NAME = "One";
    private static final String TEMPLATE_PATH = "src/main/resources/template/Blank.xls";
    private static final String OUTPUT_FILE_PATH = "src/main/resources/output/blank-mod-new.xls";

    private XlsxReadingService readingService;
    private XlsxSearchingService searchingService;
    private XlsOperationService operationService;

    @BeforeEach
    public void setup() {
        readingService = new XlsxReadingService(TEMPLATE_PATH);
        searchingService = new XlsxSearchingService();
        XlsxCellsWritingService cellsWritingService = new XlsxCellsWritingService();
        XlsxWritingService writingService = new XlsxWritingService();
        operationService = new XlsOperationService(
                readingService, writingService, cellsWritingService, searchingService
        );
    }

    @Test
    @DisplayName("Multiple Rows and Columns")
    public void testMultipleRowsAndColumns() {
        // Arrange
        List<String> values = Arrays.asList("val1", "val2", "val3", "val4", "val5");

        // Act
        WorkbookOperation workbookOperation = builder(TEMPLATE_PATH)
                .startAt(SHEET_NAME, 0, 0)
                .data(
                        new TableData<>(values, 3, 2)
                )
                .saveAs(OUTPUT_FILE_PATH)
                .build();

        operationService.execute(workbookOperation);

        // Assert
        List<SimpleEntry<Integer, Integer>> entries = List.of(
                new SimpleEntry<>(0, 0),
                new SimpleEntry<>(1, 0),
                new SimpleEntry<>(2, 0),
                new SimpleEntry<>(0, 1),
                new SimpleEntry<>(1, 1)
        );

        List<Cell> cellList = getCellsFromCoordinatesMap(entries);

        assertEquals(values.get(0), cellList.get(0).getStringCellValue());
        assertEquals(values.get(1), cellList.get(1).getStringCellValue());
        assertEquals(values.get(2), cellList.get(2).getStringCellValue());
        assertEquals(values.get(3), cellList.get(3).getStringCellValue());
        assertEquals(values.get(4), cellList.get(4).getStringCellValue());

    }


    @Test
    @DisplayName("Test writing beyond maximum rows and columns - should not write at all")
    public void testRowAndColumnWrapping() {
        // Arrange
        List<String> values = List.of("Data1", "Data2", "Data3", "Data4", "Data5", "Data6");
        WorkbookOperation operation = WorkbookOperation.builder(TEMPLATE_PATH)
                .startAt(SHEET_NAME, 0, 0)
                .data(new TableData<>(values, 2, 2)) // maxRows = 2, maxCols = 2
                .saveAs(OUTPUT_FILE_PATH)
                .build();

        // Act
        operationService.execute(operation);

        // Assert
        List<SimpleEntry<Integer, Integer>> entries = List.of(
                new SimpleEntry<>(0, 0),
                new SimpleEntry<>(1, 0),
                new SimpleEntry<>(0, 1),
                new SimpleEntry<>(1, 1),
                new SimpleEntry<>(1, 2)
        );

        List<Cell> cellList = getCellsFromCoordinatesMap(entries);

        assertEquals(values.get(0), cellList.get(0).getStringCellValue());
        assertEquals(values.get(1), cellList.get(1).getStringCellValue());
        assertEquals(values.get(2), cellList.get(2).getStringCellValue());
        assertEquals(values.get(3), cellList.get(3).getStringCellValue());
        assertNull(cellList.get(4));

    }

    @Test
    @DisplayName("Test handling of empty data list")
    public void testEmptyDataList() {
        // Arrange
        List<String> values = new ArrayList<>();
        WorkbookOperation operation = WorkbookOperation.builder(TEMPLATE_PATH)
                .startAt(SHEET_NAME, 0, 0)
                .data(new TableData<>(values, 2, 2))
                .saveAs(OUTPUT_FILE_PATH)
                .build();

        // Act & Assert
        assertDoesNotThrow(() -> operationService.execute(operation));
    }

    @Test
    @DisplayName("Test start at coordinates")
    public void testStartAtCoordinates() {
        // Arrange
        List<String> values = List.of("Data1", "Data2", "Data3", "Data4", "Data5");

        WorkbookOperation operation = WorkbookOperation.builder(TEMPLATE_PATH)
                .startAt(SHEET_NAME, 2, 2)
                .data(new TableData<>(values, 3, 2))
                .saveAs(OUTPUT_FILE_PATH)
                .build();

        // Act
        operationService.execute(operation);

        // Assert
        List<SimpleEntry<Integer, Integer>> entries = List.of(
                new SimpleEntry<>(2, 2),
                new SimpleEntry<>(3, 2),
                new SimpleEntry<>(4, 2),
                new SimpleEntry<>(2, 3),
                new SimpleEntry<>(3, 3)
        );

        List<Cell> cellList = getCellsFromCoordinatesMap(entries);

        assertEquals(values.get(0), cellList.get(0).getStringCellValue());
        assertEquals(values.get(1), cellList.get(1).getStringCellValue());
        assertEquals(values.get(2), cellList.get(2).getStringCellValue());
        assertEquals(values.get(3), cellList.get(3).getStringCellValue());
        assertEquals(values.get(4), cellList.get(4).getStringCellValue());
    }

    @Test
    @DisplayName("Test start at coordinates with cell search")
    public void testStartAtCoordinatesWithCellSearch() {
        // Arrange
        List<String> values = List.of("Data1", "Data2", "Data3", "Data4", "Data5");

        WorkbookOperation operation = WorkbookOperation.builder(TEMPLATE_PATH)
                .startAt(SHEET_NAME, "TARGET")
                .data(new TableData<>(values, 3, 2))
                .saveAs(OUTPUT_FILE_PATH)
                .build();

        // Act
        operationService.execute(operation);

        // Assert
        List<SimpleEntry<Integer, Integer>> entries = List.of(
                new SimpleEntry<>(6, 4)
        );

        List<Cell> cellList = getCellsFromCoordinatesMap(entries);

        assertEquals(values.get(0), cellList.get(0).getStringCellValue());
    }


    private List<Cell> getCellsFromCoordinatesMap(List<SimpleEntry<Integer, Integer>> entries) {
        XSSFWorkbook workbook = readingService.readFromTemplate(OUTPUT_FILE_PATH);
        SheetInfo sheetInfo = new SheetInfo(workbook, SHEET_NAME);
        CellCoordinates.SearchBuilder<Object> searchBuilder = SearchBuilder.init();

        return entries.stream()
                .map(entry -> searchingService.searchCellBySheetAndCoordinates(
                                sheetInfo,
                                searchBuilder
                                        .rowNumber(entry.getKey())
                                        .columnNumber(entry.getValue())
                                        .build()
                        ).orElse(null)
                ).toList();
    }


}
