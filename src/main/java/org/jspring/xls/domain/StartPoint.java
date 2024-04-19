package org.jspring.xls.domain;

public record StartPoint(int startRow, int startColumn) {

    public StartPoint(int startRow) {
        this(startRow, 0);
    }
}
