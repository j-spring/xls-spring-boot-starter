package org.jspring.xls.domain;

import org.apache.poi.ss.usermodel.CellType;

/**
 * Represents a wrapper for a cell value.
 *
 * @param <T> The type of the cell value.
 */
public record CellWrapper<T>(CellType cellType, T value) {
}
