package org.jspring.xls.builder;

import org.apache.poi.ss.usermodel.Cell;
import org.jspring.xls.domain.CellSearch;
import org.jspring.xls.domain.TableData;
import org.jspring.xls.enums.CellFilter;

import java.util.function.Predicate;

public class WorkbookOperation {

    private final String templatePath;
    private final String outputPath;
    private final TableData tableData;
    private final String startSheetName;
    private final int startRow;
    private final int startColumn;
    private final CellSearch<?> cellSearch;
    private final String searchFor;
    private final Predicate<Cell> filter;


    // Private constructor using Builder
    private WorkbookOperation(Builder builder) {
        this.templatePath = builder.templatePath;
        this.outputPath = builder.outputPath;
        this.tableData = builder.tableData;
        this.startSheetName = builder.startSheetName;
        this.startRow = builder.startRow;
        this.startColumn = builder.startColumn;
        this.cellSearch = builder.cellSearch;
        this.searchFor = builder.searchFor;
        this.filter = builder.filter;
    }

    // Static method to get a builder instance
    public static Builder builder(String templatePath) {
        return new Builder(templatePath);
    }

    // Getters for all properties
    public String getTemplatePath() {
        return templatePath;
    }

    public String getOutputPath() {
        return outputPath;
    }

    public TableData getTableData() {
        return tableData;
    }

    public String getStartSheetName() {
        return startSheetName;
    }

    public int getStartRow() {
        return startRow;
    }

    public int getStartColumn() {
        return startColumn;
    }

    public CellSearch<?> getCellSearch() {
        return cellSearch;
    }

    public String getSearchFor() {
        return searchFor;
    }

    public Predicate<Cell> getFilter() {
        return filter;
    }


    // Builder class
    public static class Builder {
        public CellSearch<?> cellSearch;
        public String templatePath;
        private String outputPath;
        private TableData tableData;
        private String startSheetName;
        private String searchFor;
        private Predicate<Cell> filter;
        private int startRow;
        private int startColumn;

        // Private constructor to enforce use of factory method
        private Builder(String templatePath) {
            this.templatePath = templatePath;
        }

        public Builder startAt(CellSearch<?> cellSearch) {
            this.cellSearch = cellSearch;
            return this;
        }

        public Builder startAt(String sheetName, String searchFor) {
            this.startSheetName = sheetName;
            this.searchFor = searchFor;
            this.startRow = -1;
            this.startColumn = -1;
            this.filter = CellFilter.NO_FILTER.predicate();
            return this;
        }

        public Builder startAt(String sheetName, String searchFor, Predicate<Cell> filter) {
            this.startSheetName = sheetName;
            this.searchFor = searchFor;
            this.startRow = -1;
            this.startColumn = -1;
            this.filter = filter;
            return this;
        }

        public Builder startAt(String sheetName, int row, int column) {
            this.startSheetName = sheetName;
            this.startRow = row;
            this.startColumn = column;
            return this;
        }

        public Builder saveAs(String outputPath) {
            this.outputPath = outputPath;
            return this;
        }

        public Builder data(TableData tableData) {
            this.tableData = tableData;
            return this;
        }

        public WorkbookOperation build() {
            return new WorkbookOperation(this);
        }
    }

}
