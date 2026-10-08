package org.example;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class MethodCellLayoutTest {
    private CellRangeAddress prepare(XSSFWorkbook book, String text, int rows, int width) {
        var sheet = book.createSheet();
        for (int i = 0; i < rows; i++) sheet.createRow(i).setHeightInPoints(18);
        for (int i = 0; i < 5; i++) sheet.setColumnWidth(i, width * 256);
        var font = book.createFont(); font.setFontName("한컴산뜻돋움"); font.setFontHeightInPoints((short)10);
        var style = book.createCellStyle(); style.setFont(font); style.setWrapText(true);
        var cell = sheet.getRow(0).createCell(0); cell.setCellValue(text); cell.setCellStyle(style);
        var range = new CellRangeAddress(0, rows - 1, 0, 4); sheet.addMergedRegion(range);
        return range;
    }
    @Test void twoStepsInOneIngredientRowRemainReadableAfterSave() throws Exception {
        try (var book = new XSSFWorkbook()) {
            var area = prepare(book, "① 오이는 깨끗이 씻는다.\n② 스틱 모양으로 썬다.", 1, 12);
            var sheet = book.getSheetAt(0);
            MethodCellLayout.fit(sheet, area);
            assertTrue(sheet.getRow(0).getHeightInPoints() >= 32);
            var bytes = new java.io.ByteArrayOutputStream(); book.write(bytes);
            try (var reopened = new XSSFWorkbook(new java.io.ByteArrayInputStream(bytes.toByteArray()))) {
                assertTrue(reopened.getSheetAt(0).getRow(0).getHeightInPoints() >= 32);
                assertEquals(area, reopened.getSheetAt(0).getMergedRegion(0));
                assertEquals(sheet.getRow(0).getCell(0).getStringCellValue(), reopened.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
            }
        }
    }
    @Test void enoughHeightAcrossIngredientRowsStaysUnchanged() throws Exception {
        try (var book = new XSSFWorkbook()) {
            var area = prepare(book, "① 씻는다.\r\n② 썬다.", 3, 12);
            MethodCellLayout.fit(book.getSheetAt(0), area);
            for (var row : book.getSheetAt(0)) assertEquals(18, row.getHeightInPoints());
        }
    }
    @Test void narrowColumnsWrapEvenWithoutExplicitNewline() throws Exception {
        try (var book = new XSSFWorkbook()) {
            String method = "① 오이를 깨끗하게 씻어서 먹기 좋은 스틱 모양으로 자른 뒤 제공한다.";
            var wide = prepare(book, method, 1, 16);
            var narrow = prepare(book, method, 1, 3);
            MethodCellLayout.fit(book.getSheetAt(0), wide); MethodCellLayout.fit(book.getSheetAt(1), narrow);
            assertTrue(book.getSheetAt(1).getRow(0).getHeightInPoints() > book.getSheetAt(0).getRow(0).getHeightInPoints());
        }
    }
    @Test void longMethodDistributesExtraHeightAndDoesNotShrinkOnRepeat() throws Exception {
        try (var book = new XSSFWorkbook()) {
            var area = prepare(book, String.join("\n", java.util.Collections.nCopies(12, "① 재료를 준비한다.")), 3, 12);
            var sheet = book.getSheetAt(0); MethodCellLayout.fit(sheet, area);
            float height = sheet.getRow(0).getHeightInPoints();
            assertTrue(height > 18); assertEquals(height, sheet.getRow(1).getHeightInPoints());
            MethodCellLayout.fit(sheet, area); assertEquals(height, sheet.getRow(0).getHeightInPoints());
        }
    }
    @Test void blankMethodKeepsHeight() throws Exception {
        try (var book = new XSSFWorkbook()) {
            var area = prepare(book, "", 1, 12); MethodCellLayout.fit(book.getSheetAt(0), area);
            assertEquals(18, book.getSheetAt(0).getRow(0).getHeightInPoints());
        }
    }
    @Test void oneLineKeepsOriginalHeightEvenWithLargerFont() throws Exception {
        try (var book = new XSSFWorkbook()) {
            var area = prepare(book, "① 깨끗이 씻어 제공한다.", 1, 12);
            var sheet = book.getSheetAt(0);
            book.getFontAt(sheet.getRow(0).getCell(0).getCellStyle().getFontIndex()).setFontHeightInPoints((short)11);
            MethodCellLayout.fit(sheet, area);
            assertEquals(18, sheet.getRow(0).getHeightInPoints());
        }
    }
    @Test void fittingLineDoesNotUseConservativeWrappingMargin() throws Exception {
        try (var book = new XSSFWorkbook()) {
            var area = prepare(book, "① 먹기 좋은 크기로 썰어 제공한다.", 1, 6);
            MethodCellLayout.fit(book.getSheetAt(0), area);
            assertEquals(18, book.getSheetAt(0).getRow(0).getHeightInPoints());
        }
    }
    @Test void singleLineKimchiServingNotesKeepOriginalHeight() throws Exception {
        try (var book = new XSSFWorkbook()) {
            for (String name : java.util.List.of("배추김치⑨", "깍두기⑨", "step3 배추김치⑨")) {
                var area = prepare(book, "① 배추김치는 1x1cm 이하 크기로 작게 잘라 제공한다.", 1, 3);
                var sheet = book.getSheetAt(book.getNumberOfSheets() - 1);
                MethodCellLayout.fit(sheet, area, name);
                assertEquals(18, sheet.getRow(0).getHeightInPoints());
            }
        }
    }
    @Test void multilineKimchiStillGetsEnoughHeight() throws Exception {
        try (var book = new XSSFWorkbook()) {
            var area = prepare(book, "① 먹기 좋은 크기로 자른다.\n② 제공한다.", 1, 12);
            var sheet = book.getSheetAt(0);
            MethodCellLayout.fit(sheet, area, "배추김치⑨");
            assertTrue(sheet.getRow(0).getHeightInPoints() >= 32);
        }
    }
}
