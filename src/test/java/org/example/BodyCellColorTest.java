package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class BodyCellColorTest {
    private void whiteAndBlack(Workbook book, CellStyle style) {
        assertEquals(FillPatternType.SOLID_FOREGROUND, style.getFillPattern());
        assertEquals(IndexedColors.WHITE.getIndex(), style.getFillForegroundColor());
        assertEquals(IndexedColors.BLACK.getIndex(), book.getFontAt(style.getFontIndexAsInt()).getColor());
    }

    @Test void extendedBodyAndBorderDerivedStylesKeepExplicitColors() throws Exception {
        try (var book = new XSSFWorkbook()) {
            var styles = JsonToExcel.Styles.build(book);
            for (var style : java.util.List.of(styles.bodyCenter, styles.methodMerged, styles.num1dec, styles.numGeneral))
                whiteAndBlack(book, style);
            var ctor = JsonToExcel.class.getDeclaredConstructor(java.nio.file.Path.class); ctor.setAccessible(true);
            var renderer = ctor.newInstance(java.nio.file.Path.of("."));
            var cell = book.createSheet().createRow(0).createCell(1);
            cell.setCellStyle(styles.bodyCenter);
            renderer.setBorders(cell, BorderStyle.DOUBLE);
            whiteAndBlack(book, cell.getCellStyle());
            assertEquals(FillPatternType.NO_FILL, styles.blankA.getFillPattern());
            assertNotEquals(IndexedColors.WHITE.getIndex(), styles.header.getFillForegroundColor());
        }
    }

    @Test void generalBodyPreservesSectionAndCountColors() throws Exception {
        try (var book = new XSSFWorkbook()) {
            var styles = JsonToExcelGeneral.Styles.build(book);
            for (var style : java.util.List.of(styles.bodyCenter, styles.methodMerged, styles.num1, styles.dec1))
                whiteAndBlack(book, style);
            for (var style : java.util.List.of(styles.labelAm, styles.labelLunch, styles.labelPm, styles.kidsCnt12Fill, styles.kidsCnt35Fill))
                assertNotEquals(IndexedColors.WHITE.getIndex(), style.getFillForegroundColor());
        }
    }
}
