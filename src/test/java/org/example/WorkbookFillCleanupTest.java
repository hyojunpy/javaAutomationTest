package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;
import static org.junit.jupiter.api.Assertions.*;

class WorkbookFillCleanupTest {
    private long residualColors(XSSFWorkbook book) {
        return book.getStylesSource().getFills().stream().filter(fill -> {
            var pattern = fill.getCTFill().getPatternFill();
            return pattern != null && (!pattern.isSetPatternType()
                    || pattern.getPatternType() == org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType.NONE)
                    && (pattern.isSetFgColor() || pattern.isSetBgColor());
        }).count();
    }

    @ParameterizedTest
    @ValueSource(booleans = {false, true})
    void removesBorderHelperColorsBeforeSaveWhileKeepingHeaders(boolean general) throws Exception {
        try (var book = new XSSFWorkbook()) {
            CellStyle body, header;
            if (general) {
                var styles = JsonToExcelGeneral.Styles.build(book); body = styles.bodyCenter; header = styles.header;
            } else {
                var styles = JsonToExcel.Styles.build(book); body = styles.bodyCenter; header = styles.header;
            }
            var sheet = book.createSheet();
            for (int r = 1; r <= 11; r++) sheet.createRow(r).createCell(1).setCellStyle(body);
            sheet.getRow(1).getCell(1).setCellValue("쇠고기폭찹");
            var range = new CellRangeAddress(1, 11, 1, 1); sheet.addMergedRegion(range);
            RegionUtil.setBorderBottom(BorderStyle.DOUBLE, range, sheet);
            RegionUtil.setBorderLeft(BorderStyle.DOUBLE, range, sheet);
            assertTrue(residualColors(book) > 0, "Reproduce POI's no-fill automatic color residue");
            int headerFillId = (int) ((org.apache.poi.xssf.usermodel.XSSFCellStyle) header).getCoreXf().getFillId();
            var headerFill = book.getStylesSource().getFillAt(headerFillId).getCTFill().xmlText();
            ColumnAStyle.apply(book, body);
            WorkbookFillCleanup.apply(book);
            assertEquals(0, residualColors(book));
            assertEquals(headerFill, book.getStylesSource().getFillAt(headerFillId).getCTFill().xmlText());
            var bytes = new java.io.ByteArrayOutputStream(); book.write(bytes);
            try (var saved = new XSSFWorkbook(new java.io.ByteArrayInputStream(bytes.toByteArray()))) {
                assertEquals(0, residualColors(saved));
                var cell = saved.getSheetAt(0).getRow(1).getCell(1);
                assertEquals("쇠고기폭찹", cell.getStringCellValue());
                assertEquals(FillPatternType.NO_FILL, cell.getCellStyle().getFillPattern());
                assertEquals(IndexedColors.BLACK.getIndex(), saved.getFontAt(cell.getCellStyle().getFontIndexAsInt()).getColor());
                assertEquals(1, saved.getSheetAt(0).getNumMergedRegions());
            }
        }
    }
}
