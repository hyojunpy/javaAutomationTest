package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import java.awt.font.FontRenderContext;
import java.awt.font.LineBreakMeasurer;
import java.awt.font.TextAttribute;
import java.text.AttributedString;

/** Excel does not automatically fit the height of merged, wrapped cells. */
final class MethodCellLayout {
    private static final FontRenderContext METRICS = new FontRenderContext(null, true, true);
    private MethodCellLayout() {}

    static void fit(Sheet sheet, CellRangeAddress area) {
        Row first = sheet.getRow(area.getFirstRow());
        if (first == null) return;
        Cell cell = first.getCell(area.getFirstColumn());
        if (cell == null || cell.getCellType() != CellType.STRING || cell.getStringCellValue().isBlank()) return;
        String method = cell.getStringCellValue();
        boolean singleParagraph = method.indexOf('\n') < 0 && method.indexOf('\r') < 0;
        // Only source instructions with explicit line breaks need extra height.
        if (singleParagraph) return;
        org.apache.poi.ss.usermodel.Font font = sheet.getWorkbook().getFontAt(cell.getCellStyle().getFontIndex());
        float points = font.getFontHeightInPoints();
        int style = (font.getBold() ? java.awt.Font.BOLD : 0) | (font.getItalic() ? java.awt.Font.ITALIC : 0);
        java.awt.Font measuredFont = new java.awt.Font(font.getFontName(), style, 1).deriveFont(points * 96f / 72f);
        float width = 0;
        for (int col = area.getFirstColumn(); col <= area.getLastColumn(); col++) width += sheet.getColumnWidthInPixels(col);
        // Allow for cell padding and differences between Java and Excel font rendering.
        width = Math.max(1, (width - 10) / 1.12f);
        int lines = 0;
        for (String paragraph : cell.getStringCellValue().split("\\r\\n|\\r|\\n", -1)) {
            if (paragraph.isEmpty()) { lines++; continue; }
            AttributedString text = new AttributedString(paragraph);
            text.addAttribute(TextAttribute.FONT, measuredFont);
            LineBreakMeasurer measurer = new LineBreakMeasurer(text.getIterator(), METRICS);
            while (measurer.getPosition() < paragraph.length()) { measurer.nextLayout(width); lines++; }
        }
        if (lines <= 1) return;
        float required = lines * points * 1.4f + 4;
        float existing = 0;
        for (int r = area.getFirstRow(); r <= area.getLastRow(); r++) existing += sheet.getRow(r).getHeightInPoints();
        if (required <= existing) return;
        float extraPerRow = (required - existing) / (area.getLastRow() - area.getFirstRow() + 1);
        for (int r = area.getFirstRow(); r <= area.getLastRow(); r++) {
            Row row = sheet.getRow(r);
            row.setHeightInPoints(Math.min(409, row.getHeightInPoints() + extraPerRow));
        }
    }
}
