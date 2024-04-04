package org.jspring.xls.domain;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

import static java.util.stream.StreamSupport.stream;

/**
 * Represents a search for a cell in a given sheet based on the sheet info and cell coordinates.
 *
 * @param <T> The type of the cell value.
 */
public record CellSearch<T>(
        SheetInfo sheetInfo,
        CellCoordinates<T> cellCoordinates
) {


    /**
     * Searches for a cell in the sheet based on the sheet info and cell coordinates.
     *
     * @return An Optional containing the found cell, or an empty Optional if the cell is not found.
     */
    public Optional<Cell> search() {
        List<Row> rows = stream(
                sheetInfo.getSheet().spliterator(),
                false
        ).toList();

        if (cellCoordinates.byCoordinates()) {
            return Optional.ofNullable(
                    rows.get(cellCoordinates.rowNumber())
                            .getCell(cellCoordinates.columnNumber())
            );
        }

        if (cellCoordinates.byValue()) {
            List<Cell> cellList = new ArrayList<>();
            rows.forEach(
                    row -> row.forEach(
                            cell -> addIfCellHasValue(
                                    cellList,
                                    cell,
                                    cellCoordinates.cellValue()
                            )
                    )
            );

            return cellList.stream()
                    .filter(cellCoordinates.filter())
                    .findFirst();

        }

        return Optional.empty();
    }

    private void addIfCellHasValue(List<Cell> cellList, Cell cell, T value) {

        boolean hasValue = switch (cell.getCellType()) {
            case STRING -> value instanceof String s
                    && cell.getStringCellValue().equalsIgnoreCase(s);
            case NUMERIC -> value instanceof Double || value instanceof Integer n
                    && cell.getNumericCellValue() == n;
            case BOOLEAN -> value instanceof Boolean b
                    && cell.getBooleanCellValue() == b;
            case FORMULA -> value instanceof String s
                    && cell.getCellFormula().equalsIgnoreCase(s);
            case BLANK, _NONE, ERROR -> false;
        };

        if (hasValue) {
            cellList.add(cell);
        }

    }

}
