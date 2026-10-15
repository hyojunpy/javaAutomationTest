package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class BodyCellColorTest {
    @Test void columnAColorsSurviveCleanupAndSavingOnEverySheet() throws Exception {
        try (var book = new XSSFWorkbook()) {
            var styles = JsonToExcel.Styles.build(book);
            for (String name : java.util.List.of("week", "birthday")) {
                var sheet = book.createSheet(name);
                sheet.setColumnWidth(0, 1024);
                var cell = sheet.createRow(20).createCell(0);
                cell.setCellValue("편집할 내용");
                cell.setCellStyle(styles.blankA);
            }
            ColumnAStyle.apply(book, styles.bodyCenter);
            var bytes = new java.io.ByteArrayOutputStream(); book.write(bytes);
            try (var saved = new XSSFWorkbook(new java.io.ByteArrayInputStream(bytes.toByteArray()))) {
                for (Sheet sheet : saved) {
                    noFillAndBlack(saved, sheet.getColumnStyle(0));
                    var cell = sheet.getRow(20).getCell(0);
                    noFillAndBlack(saved, cell.getCellStyle());
                    assertEquals("편집할 내용", cell.getStringCellValue());
                    assertEquals(1024, sheet.getColumnWidth(0));
                    assertEquals(BorderStyle.NONE, cell.getCellStyle().getBorderLeft());
                }
            }
        }
    }
    private void noFillAndBlack(Workbook book, CellStyle style) {
        assertEquals(FillPatternType.NO_FILL, style.getFillPattern());
        assertEquals(IndexedColors.BLACK.getIndex(), book.getFontAt(style.getFontIndexAsInt()).getColor());
    }

    @Test void extendedBodyAndBorderDerivedStylesKeepExplicitColors() throws Exception {
        try (var book = new XSSFWorkbook()) {
            var styles = JsonToExcel.Styles.build(book);
            for (var style : java.util.List.of(styles.bodyCenter, styles.methodMerged, styles.num1dec, styles.numGeneral))
                noFillAndBlack(book, style);
            var ctor = JsonToExcel.class.getDeclaredConstructor(java.nio.file.Path.class); ctor.setAccessible(true);
            var renderer = ctor.newInstance(java.nio.file.Path.of("."));
            var cell = book.createSheet().createRow(0).createCell(1);
            cell.setCellStyle(styles.bodyCenter);
            renderer.setBorders(cell, BorderStyle.DOUBLE);
            noFillAndBlack(book, cell.getCellStyle());
            assertEquals(FillPatternType.NO_FILL, styles.blankA.getFillPattern());
            assertNotEquals(IndexedColors.WHITE.getIndex(), styles.header.getFillForegroundColor());
        }
    }

    @Test void generalBodyPreservesSectionAndCountColors() throws Exception {
        try (var book = new XSSFWorkbook()) {
            var styles = JsonToExcelGeneral.Styles.build(book);
            for (var style : java.util.List.of(styles.bodyCenter, styles.methodMerged, styles.num1, styles.dec1))
                noFillAndBlack(book, style);
            for (var style : java.util.List.of(styles.labelAm, styles.labelLunch, styles.labelPm, styles.kidsCnt12Fill, styles.kidsCnt35Fill))
                assertNotEquals(IndexedColors.WHITE.getIndex(), style.getFillForegroundColor());
        }
    }
}
