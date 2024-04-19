package org.jspring.xls.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspring.xls.builder.WorkbookOperation;
import org.jspring.xls.domain.CellCoordinates;
import org.jspring.xls.domain.SheetInfo;
import org.jspring.xls.domain.StartPoint;

import java.util.Optional;

public class XlsOperationService {

    private final XlsxReadingService readingService;
    private final XlsxWritingService writingService;
    private final XlsxCellsWritingService cellsWritingService;
    private final XlsxSearchingService searchingService;

    public XlsOperationService(
            XlsxReadingService readingService,
            XlsxWritingService writingService,
            XlsxCellsWritingService cellsWritingService,
            XlsxSearchingService searchingService
    ) {
        this.readingService = readingService;
        this.writingService = writingService;
        this.cellsWritingService = cellsWritingService;
        this.searchingService = searchingService;
    }

    public void execute(WorkbookOperation operation) {
        // Read the workbook from the template
        XSSFWorkbook workbook = readingService.readFromTemplate(operation.getTemplatePath());
        SheetInfo sheetInfo = new SheetInfo(
                workbook, operation.getStartSheetName()
        );

        Optional<Cell> cell = searchingService.searchCellBySheetAndCoordinates(
                sheetInfo,
                new CellCoordinates<>(
                        operation.getStartRow(),
                        operation.getStartColumn(),
                        operation.getSearchFor(),
                        operation.getFilter()
                )
        );

        // Perform the table write operation
        cellsWritingService.writeTopToBottom(
                sheetInfo,
                new StartPoint(
                        cell.map(Cell::getRowIndex)
                                .orElseGet(operation::getStartRow),
                        cell.map(Cell::getColumnIndex)
                                .orElseGet(operation::getStartColumn)
                ),
                operation.getTableData()
        );

        // Save the workbook
        writingService.writeFile(workbook, operation.getOutputPath());
    }

}
