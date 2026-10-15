package org.example;

import org.apache.poi.ss.usermodel.*;

/** Apply the requested no-fill body style to column A after template cleanup. */
final class ColumnAStyle {
    private ColumnAStyle() {}

    static void apply(Workbook workbook, CellStyle bodyStyle) {
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(bodyStyle);
        style.setFillPattern(FillPatternType.NO_FILL);
        style.setBorderTop(BorderStyle.NONE);
        style.setBorderBottom(BorderStyle.NONE);
        style.setBorderLeft(BorderStyle.NONE);
        style.setBorderRight(BorderStyle.NONE);
        for (Sheet sheet : workbook) {
            sheet.setDefaultColumnStyle(0, style);
            for (Row row : sheet) {
                Cell cell = row.getCell(0);
                if (cell != null) cell.setCellStyle(style);
            }
        }
    }
}
