package org.jspring.xls.integration;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspring.xls.domain.CellCoordinates;
import org.jspring.xls.domain.CellSearch;
import org.jspring.xls.domain.SheetInfo;
import org.jspring.xls.service.XlsxReadingService;
import org.jspring.xls.service.XlsxSearchingService;
import org.jspring.xls.service.XlsxTableService;
import org.jspring.xls.service.XlsxWritingService;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.List;
import java.util.Optional;

import static org.junit.jupiter.api.Assertions.assertNotNull;

@SpringBootTest
class XlsxReadAndWriteTest {

    private static final String SHEET_NAME = "One";
    private static final String OUTPUT_FILE_PATH = "src/main/resources/output/blank-mod.xls";

    @Autowired
    private XlsxReadingService xlsxReadingService;
    @Autowired
    private XlsxWritingService xlsxWritingService;
    @Autowired
    private XlsxTableService tableService;
    @Autowired
    private XlsxSearchingService searchingService;


    @Test
    @DisplayName("Read Xlsx from template")
    void readFromTemplate() {
        XSSFWorkbook workbook = xlsxReadingService.readFromTemplate();
        SheetInfo sheetInfo = new SheetInfo(workbook, SHEET_NAME);
        assertNotNull(sheetInfo);
    }

    @Test
    @DisplayName("Write Xlsx to file")
    void writeToFile() {
        XSSFWorkbook workbook = xlsxReadingService.readFromTemplate();
        SheetInfo sheetInfo = new SheetInfo(workbook, SHEET_NAME);
        assertNotNull(sheetInfo);
        xlsxWritingService.writeFile(workbook, OUTPUT_FILE_PATH);

        XSSFWorkbook workbookMod = xlsxReadingService.readFromTemplate(OUTPUT_FILE_PATH);
        SheetInfo sheetInfoMod = new SheetInfo(workbookMod, SHEET_NAME);
        assertNotNull(sheetInfoMod);
    }

    @Test
    @DisplayName("Write and read Xlsx from byte array")
    void writeAndReadFromByteArray() {
        XSSFWorkbook workbookFromTemplate = xlsxReadingService.readFromTemplate();
        byte[] bytes = xlsxWritingService.writeAsByteArray(workbookFromTemplate);
        assertNotNull(bytes);

        XSSFWorkbook workbookFromByteArray = xlsxReadingService.readFromByteArray(bytes);
        SheetInfo sheetInfoFromByteArray = new SheetInfo(workbookFromByteArray, SHEET_NAME);
        assertNotNull(sheetInfoFromByteArray);
    }


    @Test
    @DisplayName("Add table to existing file")
    void addTableToExistingFile() {
        XSSFWorkbook workbook = xlsxReadingService.readFromTemplate();
        SheetInfo sheetInfo = new SheetInfo(workbook, SHEET_NAME);
        assertNotNull(sheetInfo);

        // add table
        tableService.populateWorksheetWithData(
                new CellSearch<>(
                        new SheetInfo(workbook, SHEET_NAME),
                        CellCoordinates.SearchBuilder.init()
                                .address("D6")
                                .build()
                ),
                4,
                List.of(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16)
        );

        // write new file
        xlsxWritingService.writeFile(workbook, OUTPUT_FILE_PATH);

        // read file just created
        XSSFWorkbook workbookMod = xlsxReadingService.readFromTemplate(OUTPUT_FILE_PATH);
        SheetInfo sheetInfoMod = new SheetInfo(workbookMod, SHEET_NAME);
        assertNotNull(sheetInfoMod);
    }

    @Test
    @DisplayName("Add table of strings to existing file")
    void addTableOfStringsToExistingFile() {
        XSSFWorkbook workbook = xlsxReadingService.readFromTemplate();
        SheetInfo sheetInfo = new SheetInfo(workbook, SHEET_NAME);
        assertNotNull(sheetInfo);

        // add table
        tableService.populateWorksheetWithData(
                new CellSearch<>(
                        new SheetInfo(workbook, SHEET_NAME),
                        CellCoordinates.SearchBuilder.init()
                                .address("D6")
                                .build()
                ),
                3,
                List.of("1","2","3","4","5","6","7")
        );

        // write new file
        xlsxWritingService.writeFile(workbook, OUTPUT_FILE_PATH);

        // read file just created
        XSSFWorkbook workbookMod = xlsxReadingService.readFromTemplate(OUTPUT_FILE_PATH);
        SheetInfo sheetInfoMod = new SheetInfo(workbookMod, SHEET_NAME);
        assertNotNull(sheetInfoMod);
    }

    @Test
    @DisplayName("Add table to existing file using coordinates")
    void addTableToExistingFileUsingCoordinates() {
        XSSFWorkbook workbook = xlsxReadingService.readFromTemplate();
        SheetInfo sheetInfo = new SheetInfo(workbook, SHEET_NAME);
        assertNotNull(sheetInfo);

        Optional<Cell> cellX = searchingService.searchCellBySheetAndCoordinates(
                sheetInfo,
                CellCoordinates.SearchBuilder.init()
                        .cellValue(202301)
                        .build()
        );

        Optional<Cell> cellY = searchingService.searchCellBySheetAndCoordinates(
                sheetInfo,
                CellCoordinates.SearchBuilder.init()
                        .cellValue("202302")
                        .firstColumn()
                        .build()
        );

        Cell cellByCoordinates = searchingService.getCellByCoordinates(cellX.get(), cellY.get());

        // add table
        tableService.populateWorksheetWithData(
                new CellSearch<>(
                        new SheetInfo(workbook, SHEET_NAME),
                        CellCoordinates.SearchBuilder.init()
                                .rowNumber(cellByCoordinates.getRowIndex())
                                .columnNumber(cellByCoordinates.getColumnIndex())
                                .build()
                ),
                3,
                List.of("1","2","3","4","5","6","7")
        );

        // write new file
        xlsxWritingService.writeFile(workbook, OUTPUT_FILE_PATH);

        // read file just created
        XSSFWorkbook workbookMod = xlsxReadingService.readFromTemplate(OUTPUT_FILE_PATH);
        SheetInfo sheetInfoMod = new SheetInfo(workbookMod, SHEET_NAME);
        assertNotNull(sheetInfoMod);
    }

}