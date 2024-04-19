package org.jspring.xls.domain;

import java.util.List;

public record TableData<T>(List<T> values, int maxRows, int maxCols) {
}
