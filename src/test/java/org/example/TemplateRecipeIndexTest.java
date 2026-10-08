package org.example;

import java.nio.file.Path;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class TemplateRecipeIndexTest {
    private JsonToExcelGeneral renderer() throws Exception {
        var constructor = JsonToExcelGeneral.class.getDeclaredConstructor(Path.class);
        constructor.setAccessible(true);
        return constructor.newInstance(Path.of("."));
    }

    private void recipe(org.apache.poi.ss.usermodel.Sheet sheet, int rowIndex, String ingredient) {
        var row = sheet.createRow(rowIndex);
        row.createCell(2).setCellValue("step1 함박스테이크①&소스⑤");
        row.createCell(3).setCellValue(ingredient);
        row.createCell(5).setCellValue(10);
    }

    @Test void preservesNewestSheetThenFirstBlockAndMergedIngredients() throws Exception {
        try (var book = new XSSFWorkbook()) {
            recipe(book.createSheet("old"), 0, "old ingredient");
            var latest = book.createSheet("new");
            recipe(latest, 1, "first ingredient");
            latest.createRow(2).createCell(3).setCellValue("second ingredient");
            latest.addMergedRegion(new CellRangeAddress(1, 2, 2, 2));
            recipe(latest, 4, "duplicate ingredient");
            var renderer = renderer();
            var cols = renderer.colsOf(JsonToExcelGeneral.TplKind.P35);
            for (int i = 0; i < 3; i++) {
                var block = renderer.findBlockInTemplateExact(book, "함박스테이크&소스", cols).orElseThrow();
                assertEquals(java.util.List.of("first ingredient", "second ingredient"),
                        block.items.stream().map(item -> item.ingredient).toList());
                assertTrue(renderer.findBlockInTemplateExact(book, "없는메뉴", cols).isEmpty());
            }
        }
    }

    @Test void separatesWorkbooksAndKeepsOutputMergeLookupsLive() throws Exception {
        var renderer = renderer();
        var cols = renderer.colsOf(JsonToExcelGeneral.TplKind.P35);
        try (var one = new XSSFWorkbook(); var two = new XSSFWorkbook(); var output = new XSSFWorkbook()) {
            recipe(one.createSheet(), 0, "one"); recipe(two.createSheet(), 0, "two");
            assertEquals("one", renderer.findBlockInTemplateExact(one, "함박스테이크&소스", cols).orElseThrow().items.get(0).ingredient);
            assertEquals("two", renderer.findBlockInTemplateExact(two, "함박스테이크&소스", cols).orElseThrow().items.get(0).ingredient);
            var sheet = output.createSheet();
            assertNull(renderer.findMergedRange(sheet, 0, 0));
            sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 0));
            assertNotNull(renderer.findMergedRange(sheet, 0, 0));
        }
    }
}
