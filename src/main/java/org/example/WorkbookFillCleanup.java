package org.example;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType;

/** Must run last: POI's border helpers can reintroduce automatic colors into no-fill styles. */
final class WorkbookFillCleanup {
    private WorkbookFillCleanup() {}

    static void apply(Workbook workbook) {
        if (!(workbook instanceof XSSFWorkbook xssf)) return;
        for (var fill : xssf.getStylesSource().getFills()) {
            var xml = fill.getCTFill();
            if (!xml.isSetPatternFill()) continue;
            var pattern = xml.getPatternFill();
            if (pattern.isSetPatternType() && pattern.getPatternType() != STPatternType.NONE) continue;
            if (pattern.isSetFgColor()) pattern.unsetFgColor();
            if (pattern.isSetBgColor()) pattern.unsetBgColor();
            pattern.setPatternType(STPatternType.NONE);
        }
    }
}
