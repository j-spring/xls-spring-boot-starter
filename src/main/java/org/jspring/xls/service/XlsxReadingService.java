package org.jspring.xls.service;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.IOException;


/**
 * The XlsxReadingService class provides functionality for reading data from an XLSX file.
 */
public class XlsxReadingService {

    private final String templatePath;
    public XlsxReadingService(String templatePath) {
        this.templatePath = templatePath;
    }

    /**
     * Reads an XSSFWorkbook from a template file.
     *
     * @return The XSSFWorkbook read from the template file.
     * @throws RuntimeException if an IOException occurs while reading the template file.
     */
    public XSSFWorkbook readFromTemplate() {
        try (FileInputStream inputStream = new FileInputStream(templatePath)) {
            return new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Reads an {@link XSSFWorkbook} from a template file.
     *
     * @param templatePath The path of the template file.
     * @return The {@link XSSFWorkbook} read from the template file.
     * @throws RuntimeException if an {@link IOException} occurs while reading the template file.
     */
    public XSSFWorkbook readFromTemplate(String templatePath) {
        try (FileInputStream inputStream = new FileInputStream(templatePath)) {
            return new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Reads an XSSFWorkbook from a byte array.
     *
     * @param fileContent The byte array containing the workbook data.
     * @return The XSSFWorkbook read from the byte array.
     * @throws RuntimeException if an IOException occurs while reading the byte array.
     */
    public XSSFWorkbook readFromByteArray(byte[] fileContent) {
        try (ByteArrayInputStream inputStream = new ByteArrayInputStream(fileContent)) {
            return new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Searches for a cell in a given sheet based on the sheet info and cell coordinates.
     *
     * @param sheetInfo        The sheet info containing the workbook and sheet name.
     * @param cellCoordinates  The cell coordinates containing the row number, column number, cell value, and filter.
     * @param <T>              The type of the cell value.
     * @return An optional containing the found cell, or an empty optional if the cell is not found.
     *//*
    public <T> Optional<Cell> searchCellBySheetAndCoordinates(
            SheetInfo sheetInfo,
            CellCoordinates<T> cellCoordinates
    ) {

        return new CellSearch<>(sheetInfo, cellCoordinates)
                .search();

    }*/

    /**
     * Retrieves the cell at the specified coordinates.
     *
     * @param cellX The X coordinate of the cell.
     * @param cellY The Y coordinate of the cell.
     * @return The cell at the specified coordinates.
     *//*
    public Cell getCellByCoordinates(Cell cellX, Cell cellY) {
        return getCellByCoordinates(cellY, cellX.getColumnIndex());
    }

    *//**
     * Retrieves the cell to the right of the given cell.
     *
     * @param cell The cell to the left of the desired cell.
     * @return The cell to the right of the given cell, or null if no cell is found.
     *//*
    public Cell getRightCell(Cell cell) {
        return getCellByCoordinates(cell, cell.getColumnIndex() + 1);
    }

    *//**
     * Retrieves the cell to the left of the given cell.
     *
     * @param cell The cell to the right of the desired cell.
     * @return The cell to the left of the given cell, or null if no cell is found.
     *//*
    public Cell getLeftCell(Cell cell) {
        return getCellByCoordinates(cell, cell.getColumnIndex() - 1);
    }

    private Cell getCellByCoordinates(Cell cellY, int columnIndex) {

        return cellY.getRow().getCell(
                columnIndex
        );
    }

    *//**
     * Reads the value of a given cell and returns it wrapped in a CellWrapper object.
     *
     * @param cell The cell to read the value from.
     * @return The value of the cell wrapped in a CellWrapper object.
     *//*
    public CellWrapper<?> readCellValue(Cell cell) {

        return switch (cell.getCellType()) {
            case STRING -> new CellWrapper<>(cell.getCellType(), cell.getStringCellValue());
            case NUMERIC -> new CellWrapper<>(cell.getCellType(), cell.getNumericCellValue());
            case BOOLEAN -> new CellWrapper<>(cell.getCellType(), cell.getBooleanCellValue());
            case FORMULA -> new CellWrapper<>(cell.getCellType(), cell.getCellFormula());
            case BLANK, _NONE, ERROR -> new CellWrapper<>(cell.getCellType(), "Unhandled");
        };
    }*/
}