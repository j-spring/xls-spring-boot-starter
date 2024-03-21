package org.jspring.xls.domain;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Represents the information of a sheet in a workbook.
 *
 * @param workbook  The XSSFWorkbook containing the sheet.
 * @param sheetName The name of the sheet.
 */
public record SheetInfo(
        XSSFWorkbook workbook,
        String sheetName
) {

    /**
     * Retrieves the sheet with the given name from the workbook.
     *
     * @return The XSSFSheet object corresponding to the specified sheet name.
     */
    public XSSFSheet getSheet() {
        return workbook().getSheet(sheetName());
    }

}
