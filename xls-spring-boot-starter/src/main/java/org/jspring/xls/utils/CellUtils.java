package org.jspring.xls.utils;

import org.apache.poi.ss.usermodel.Cell;

public class CellUtils {

    /**
     * Writes a value to a cell in an XLSX workbook.
     *
     * @param cell  The cell to write the value to.
     * @param value The value to write to the cell.
     * @throws IllegalStateException If the value type is unexpected.
     */
    public static void writeValue(
            Cell cell, Object value) {

        switch (value) {
            case String stringVal -> cell.setCellValue(stringVal);
            case Double doubleVal -> cell.setCellValue(doubleVal);
            case Boolean booleanVal -> cell.setCellValue(booleanVal);
            case Integer intVal -> cell.setCellValue(intVal);
            default -> throw new IllegalStateException("Unexpected value type: " + value.getClass().getName());
        }

    }

}
