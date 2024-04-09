package org.jspring.xls.integration;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspring.xls.config.XlsConfiguration;
import org.jspring.xls.domain.SheetInfo;
import org.jspring.xls.service.XlsxReadingService;
import org.jspring.xls.service.XlsxSearchingService;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.Optional;

import static org.jspring.xls.domain.CellCoordinates.SearchBuilder;

@SpringBootTest(classes = XlsConfiguration.class)
class XlsxReadingServiceTest {

    private static final String SHEET_NAME = "One";

    @Autowired
    private XlsxReadingService xlsxReadingService;
    @Autowired
    private XlsxSearchingService xlsxSearchingService;

    private SheetInfo sheetInfo;


    @BeforeEach
    public void setUp() {
        XSSFWorkbook workbook = xlsxReadingService.readFromTemplate();
        sheetInfo = new SheetInfo(workbook, SHEET_NAME);
    }

    @Test
    @DisplayName("Search cell by value")
    void testXlsxTargetCell() {

        Optional<Cell> cell = xlsxSearchingService.searchCellBySheetAndCoordinates(
                sheetInfo,
                SearchBuilder.init()
                        .cellValue("Voce :")
                        .build()
        );

        if (cell.isPresent()) {
            System.out.println("found!!!");
        }

    }

    @Test
    @DisplayName("Search cell by value in first column")
    void testXlsxTargetCellFirstCol() {

        Optional<Cell> cellY = xlsxSearchingService.searchCellBySheetAndCoordinates(
                sheetInfo,
                SearchBuilder.init()
                        .cellValue(202311.0)
                        .firstColumn()
                        .build()
        );

        if (cellY.isPresent()) {
            System.out.println("found");
        }

    }


    @Test
    @DisplayName("Search two cells by value")
    void testXlsxTargetCellAndWriteWithNamedCoordinates() {

        Optional<Cell> cellX = xlsxSearchingService.searchCellBySheetAndCoordinates(
                sheetInfo,
                SearchBuilder.init()
                        .cellValue("Quota")
                        .build()
        );

        Optional<Cell> cellY = xlsxSearchingService.searchCellBySheetAndCoordinates(
                sheetInfo,
                SearchBuilder.init()
                        .cellValue(202311)
                        .firstColumn()
                        .build()
        );

        if (cellX.isPresent() && cellY.isPresent()) {
            System.out.println("found coordinates: x -> " + cellX.get().getColumnIndex() + " y -> " + cellY.get().getRowIndex());
        }

    }

    @Test
    @DisplayName("Get cell value by coordinates")
    void testXlsxTargetCellAndWriteWithCoordinates() {

        xlsxSearchingService.searchCellBySheetAndCoordinates(
                        sheetInfo,
                        SearchBuilder.init()
                                .rowNumber(2)
                                .columnNumber(0)
                                .build()
                )
                .ifPresent(
                        cell -> System.out.println("found coordinates value: " + xlsxSearchingService.readCellValue(cell).value())
                );

    }

    @Test
    @DisplayName("Get cell value by cell address")
    void testXlsxTargetCellAndWriteWithAddress() {

        xlsxSearchingService.searchCellBySheetAndCoordinates(
                        sheetInfo,
                        SearchBuilder.init()
                                .address("A2")
                                .build()
                )
                .ifPresent(
                        cell -> System.out.println("found coordinates address value: " + xlsxSearchingService.readCellValue(cell).value())
                );

    }


}