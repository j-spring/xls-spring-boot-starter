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
        SheetInfo sheetInfo = new SheetInfo(workbook, "Risconti");
        assertNotNull(sheetInfo);
    }

    @Test
    @DisplayName("Write Xlsx to file")
    void writeToFile() {
        XSSFWorkbook workbook = xlsxReadingService.readFromTemplate();
        SheetInfo sheetInfo = new SheetInfo(workbook, "Risconti");
        assertNotNull(sheetInfo);
        xlsxWritingService.writeFile(workbook, "src/main/resources/xlsx/output/risconti-mod.xls");

        XSSFWorkbook workbookMod = xlsxReadingService.readFromTemplate("src/main/resources/xlsx/output/risconti-mod.xls");
        SheetInfo sheetInfoMod = new SheetInfo(workbookMod, "Risconti");
        assertNotNull(sheetInfoMod);
    }

    @Test
    @DisplayName("Write and read Xlsx from byte array")
    void writeAndReadFromByteArray() {
        XSSFWorkbook workbookFromTemplate = xlsxReadingService.readFromTemplate();
        byte[] bytes = xlsxWritingService.writeAsByteArray(workbookFromTemplate);
        assertNotNull(bytes);

        XSSFWorkbook workbookFromByteArray = xlsxReadingService.readFromByteArray(bytes);
        SheetInfo sheetInfoFromByteArray = new SheetInfo(workbookFromByteArray, "Risconti");
        assertNotNull(sheetInfoFromByteArray);
    }


    @Test
    @DisplayName("Add table to existing file")
    void addTableToExistingFile() {
        XSSFWorkbook workbook = xlsxReadingService.readFromTemplate();
        SheetInfo sheetInfo = new SheetInfo(workbook, "Risconti");
        assertNotNull(sheetInfo);

        // add table
        tableService.populateWorksheetWithData(
                new CellSearch<>(
                        new SheetInfo(workbook, "Risconti"),
                        CellCoordinates.SearchBuilder.init()
                                //.rowNumber(4)
                                //.columnNumber(3)
                                .address("D6")
                                .build()
                ),
                4,
                List.of(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16)
        );

        // write new file
        xlsxWritingService.writeFile(workbook, "src/main/resources/xlsx/output/risconti-mod.xls");

        // read file just created
        XSSFWorkbook workbookMod = xlsxReadingService.readFromTemplate("src/main/resources/xlsx/output/risconti-mod.xls");
        SheetInfo sheetInfoMod = new SheetInfo(workbookMod, "Risconti");
        assertNotNull(sheetInfoMod);
    }

    @Test
    @DisplayName("Add table of strings to existing file")
    void addTableOfStringsToExistingFile() {
        XSSFWorkbook workbook = xlsxReadingService.readFromTemplate();
        SheetInfo sheetInfo = new SheetInfo(workbook, "Risconti");
        assertNotNull(sheetInfo);

        // add table
        tableService.populateWorksheetWithData(
                new CellSearch<>(
                        new SheetInfo(workbook, "Risconti"),
                        CellCoordinates.SearchBuilder.init()
                                .address("D6")
                                .build()
                ),
                3,
                List.of("1","2","3","4","5","6","7")
        );

        // write new file
        xlsxWritingService.writeFile(workbook, "src/main/resources/xlsx/output/risconti-mod.xls");

        // read file just created
        XSSFWorkbook workbookMod = xlsxReadingService.readFromTemplate("src/main/resources/xlsx/output/risconti-mod.xls");
        SheetInfo sheetInfoMod = new SheetInfo(workbookMod, "Risconti");
        assertNotNull(sheetInfoMod);
    }

    @Test
    @DisplayName("Add table to existing file using coordinates")
    void addTableToExistingFileUsingCoordinates() {
        XSSFWorkbook workbook = xlsxReadingService.readFromTemplate();
        SheetInfo sheetInfo = new SheetInfo(workbook, "Risconti");
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
                        new SheetInfo(workbook, "Risconti"),
                        CellCoordinates.SearchBuilder.init()
                                .rowNumber(cellByCoordinates.getRowIndex())
                                .columnNumber(cellByCoordinates.getColumnIndex())
                                .build()
                ),
                3,
                List.of("1","2","3","4","5","6","7")
        );

        // write new file
        xlsxWritingService.writeFile(workbook, "src/main/resources/xlsx/output/risconti-mod.xls");

        // read file just created
        XSSFWorkbook workbookMod = xlsxReadingService.readFromTemplate("src/main/resources/xlsx/output/risconti-mod.xls");
        SheetInfo sheetInfoMod = new SheetInfo(workbookMod, "Risconti");
        assertNotNull(sheetInfoMod);
    }

}