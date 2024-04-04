package org.jspring.xls.domain;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellReference;
import org.jspring.xls.enums.CellFilter;

import java.util.function.Predicate;

/**
 * Represents the coordinates of a cell, including the row number, column number, cell value, and filter.
 *
 * @param <T> The type of the cell value.
 */
public record CellCoordinates<T>(
        int rowNumber,
        int columnNumber,
        T cellValue,
        Predicate<Cell> filter
) {

    /**
     * Checks if the cell coordinates are valid by verifying that both the row number and column number are not -1.
     *
     * @return True if the cell coordinates are valid, false otherwise.
     */
    public boolean byCoordinates() {
        return rowNumber != -1 && columnNumber != -1;
    }

    /**
     * Checks if the cell value is not null.
     *
     * @return True if the cell value is not null, false otherwise.
     */
    public boolean byValue() {
        return cellValue != null;
    }

    /**
     * A builder class for creating instances of CellCoordinates.
     *
     * @param <T> The type of the cell value.
     */
    public static class SearchBuilder<T> {
        private int rowNumber = -1;
        private int columnNumber = -1;
        private T cellValue;
        private Predicate<Cell> filter = CellFilter.NO_FILTER.predicate();

        /**
         * Initializes a new instance of the SearchBuilder class.
         *
         * @param <T> The type of the cell value.
         * @return A new instance of the SearchBuilder class.
         */
        public static <T> SearchBuilder<T> init() {
            return new SearchBuilder<>();
        }

        /**
         * Sets the row number for the cell coordinates.
         *
         * @param rowNumber The row number to set.
         * @return The SearchBuilder instance.
         */
        public SearchBuilder<T> rowNumber(int rowNumber) {
            this.rowNumber = rowNumber;
            return this;
        }

        /**
         * Sets the column number for the cell coordinates.
         *
         * @param columnNumber The column number to set.
         * @return The SearchBuilder instance.
         */
        public SearchBuilder<T> columnNumber(int columnNumber) {
            this.columnNumber = columnNumber;
            return this;
        }

        /**
         * Sets the cell value for the cell coordinates.
         *
         * @param cellValue The cell value to set.
         * @return The SearchBuilder instance.
         */
        public SearchBuilder<T> cellValue(T cellValue) {
            this.cellValue = cellValue;
            return this;
        }

        /**
         * Sets the filter for the cell coordinates.
         *
         * @param filter The predicate used to filter the cells.
         * @return The SearchBuilder instance.
         */
        public SearchBuilder<T> filter(Predicate<Cell> filter) {
            this.filter = filter;
            return this;
        }

        /**
         * Sets the filter for the cell coordinates to the first column.
         *
         * @return The SearchBuilder instance.
         */
        public SearchBuilder<T> firstColumn() {
            this.filter = CellFilter.FIRST_COLUMN.predicate();
            return this;
        }

        /**
         * Sets the filter for the cell coordinates to the first row.
         *
         * @return The SearchBuilder instance with the first row filter set.
         */
        public SearchBuilder<T> firstRow() {
            this.filter = CellFilter.FIRST_ROW.predicate();
            return this;
        }

        /**
         * Sets the address for the cell coordinates.
         *
         * @param address The address to set.
         * @return The SearchBuilder instance.
         */
        public SearchBuilder<T> address(String address) {
            CellReference reference = new CellReference(address);
            this.rowNumber = reference.getRow();
            this.columnNumber = reference.getCol();
            return this;
        }

        /**
         * Builds a new instance of CellCoordinates using the provided values.
         *
         * @return A new instance of CellCoordinates.
         */
        public CellCoordinates<T> build() {
            return new CellCoordinates<>(rowNumber, columnNumber, cellValue, filter);
        }

    }

}
