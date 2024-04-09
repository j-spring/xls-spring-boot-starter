package org.jspring.xls.integration;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspring.xls.config.XlsConfiguration;
import org.jspring.xls.domain.SheetInfo;
import org.jspring.xls.service.XlsxReadingService;
import org.jspring.xls.service.XlsxSearchingService;
import org.jspring.xls.service.XlsxWritingService;
import org.jspring.xls.utils.CellUtils;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.Optional;

import static org.jspring.xls.domain.CellCoordinates.SearchBuilder;

@SpringBootTest(classes = XlsConfiguration.class)
class XlsxWritingServiceTest {

    private static final String SHEET_NAME = "One";
    private static final String OUTPUT_FILE_PATH = "src/main/resources/output/blank-mod.xls";

    @Autowired
    private XlsxReadingService xlsxReadingService;
    @Autowired
    private XlsxWritingService xlsxWritingService;
    @Autowired
    private XlsxSearchingService xlsxSearchingService;

    private XSSFWorkbook workbook;
    private SheetInfo sheetInfo;


    @BeforeEach
    public void setUp() {
        workbook = xlsxReadingService.readFromTemplate();
        sheetInfo = new SheetInfo(workbook, SHEET_NAME);
    }


    @Test
    @DisplayName("Write to the right of cell searched by value")
    void testXlsxTargetCellAndWriteToTheRight() {

        Optional<Cell> cell = xlsxSearchingService.searchCellBySheetAndCoordinates(
                sheetInfo,
                SearchBuilder.init()
                        .cellValue("Voce :")
                        .build()
        );

        if (cell.isPresent()) {
            System.out.println("found!!!");
            CellUtils.writeValue(
                    xlsxSearchingService.getRightCell(cell.get()),
                    "voce1234"
            );

            xlsxWritingService.writeFile(workbook, OUTPUT_FILE_PATH);
        }

    }


    @Test
    @DisplayName("Write at coordinates of two cells searched by value")
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
            CellUtils.writeValue(
                    xlsxSearchingService.getCellByCoordinates(cellX.get(), cellY.get()),
                    12.34
            );

            xlsxWritingService.writeFile(workbook, OUTPUT_FILE_PATH);
        }

    }


    @Test
    @DisplayName("Evaluate formula to the right of cell searched by value")
    void testXlsxEvaluateFormulas() {

        Optional<Cell> cell = xlsxSearchingService.searchCellBySheetAndCoordinates(
                sheetInfo,
                SearchBuilder.init()
                        .cellValue("TOTALI ")
                        .firstColumn()
                        .build()
        );

        if (cell.isPresent()) {
            Cell cellRight = cell.get().getRow().getCell(cell.get().getColumnIndex() + 1);
            cellRight.setCellFormula("SUM(B6:B17)");
            XSSFFormulaEvaluator formulaEvaluator =
                    workbook.getCreationHelper().createFormulaEvaluator();
            formulaEvaluator.evaluate(cellRight);
        }

        xlsxWritingService.writeFile(workbook, OUTPUT_FILE_PATH);

    }


}