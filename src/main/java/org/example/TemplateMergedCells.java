package org.example;

import java.util.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/** Snapshot for read-only template sheets; never used for mutable output sheets. */
final class TemplateMergedCells {
    private final Map<Integer, List<CellRangeAddress>> rows = new HashMap<>();

    TemplateMergedCells(Sheet sheet) {
        for (CellRangeAddress range : sheet.getMergedRegions()) {
            // Only rows that can contain template data need a lookup entry.
            for (int row = range.getFirstRow(); row <= Math.min(range.getLastRow(), sheet.getLastRowNum()); row++) {
                rows.computeIfAbsent(row, ignored -> new ArrayList<>()).add(range);
            }
        }
    }

    CellRangeAddress find(int row, int column) {
        for (CellRangeAddress range : rows.getOrDefault(row, List.of())) {
            if (range.isInRange(row, column)) return range;
        }
        return null;
    }
}
