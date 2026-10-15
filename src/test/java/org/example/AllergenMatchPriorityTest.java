package org.example;

import java.nio.file.Path;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;
import static org.junit.jupiter.api.Assertions.*;

class AllergenMatchPriorityTest {
    private void recipe(Workbook book, boolean extended, String sheetName, String menu, String ingredient) {
        var row = book.createSheet(sheetName).createRow(0);
        row.createCell(extended ? 1 : 2).setCellValue(menu);
        row.createCell(extended ? 2 : 3).setCellValue(ingredient);
        row.createCell(extended ? 4 : 5).setCellValue(10);
    }

    private String find(boolean extended, Workbook recent, Workbook old, String query) throws Exception {
        if (extended) {
            var ctor = JsonToExcel.class.getDeclaredConstructor(Path.class); ctor.setAccessible(true);
            var result = ctor.newInstance(Path.of(".")).findBlockInTemplateExactPreferNew(recent, old, query);
            return result.found ? result.block.items.get(0).ingredient : null;
        }
        var ctor = JsonToExcelGeneral.class.getDeclaredConstructor(Path.class); ctor.setAccessible(true);
        var renderer = ctor.newInstance(Path.of("."));
        return renderer.findBlockInTemplateExactPreferNew(recent, old, query,
                renderer.colsOf(JsonToExcelGeneral.TplKind.P35)).map(b -> b.items.get(0).ingredient).orElse(null);
    }

    @ParameterizedTest @ValueSource(booleans={false,true})
    void exactAllergensWinOverNewerSheetWithoutThoseNumbers(boolean extended) throws Exception {
        try (var book = new XSSFWorkbook()) {
            recipe(book, extended, "26.2월 넷째주", "숙주나물무침⑤⑥", "간장");
            recipe(book, extended, "26. 7월 셋째주", "숙주나물무침", "소금");
            assertEquals("간장", find(extended, book, null, "숙주나물무침⑤⑥"));
            assertEquals("소금", find(extended, book, null, "숙주나물무침"));
        }
    }

    @ParameterizedTest @ValueSource(booleans={false,true})
    void exactOldTemplateWinsOverNormalizedNewTemplate(boolean extended) throws Exception {
        try (var recent = new XSSFWorkbook(); var old = new XSSFWorkbook()) {
            recipe(recent, extended, "new", "숙주나물무침", "new normalized");
            recipe(old, extended, "old", "숙주나물무침⑤⑥", "old exact");
            assertEquals("old exact", find(extended, recent, old, "숙주나물무침⑤⑥"));
        }
    }

    @ParameterizedTest @ValueSource(booleans={false,true})
    void matchingNumbersKeepNewTemplateAndNewestSheetPriority(boolean extended) throws Exception {
        try (var recent = new XSSFWorkbook(); var old = new XSSFWorkbook()) {
            recipe(recent, extended, "earlier", "숙주나물무침⑤⑥", "earlier exact");
            recipe(recent, extended, "latest", "숙주나물무침⑤⑥", "latest exact");
            recipe(old, extended, "old", "숙주나물무침⑤⑥", "old exact");
            assertEquals("latest exact", find(extended, recent, old, "숙주나물무침⑤⑥"));
        }
    }

    @ParameterizedTest @ValueSource(booleans={false,true})
    void fallsBackOnlyWhenNoTemplateHasMatchingNumbers(boolean extended) throws Exception {
        try (var recent = new XSSFWorkbook(); var old = new XSSFWorkbook()) {
            recipe(recent, extended, "earlier", "숙주나물무침⑤", "earlier normalized");
            recipe(recent, extended, "latest", "숙주나물무침⑥", "latest normalized");
            recipe(old, extended, "old", "숙주나물무침⑦", "old normalized");
            assertEquals("latest normalized", find(extended, recent, old, "숙주나물무침①②"));
            assertNull(find(extended, recent, old, "없는메뉴⑤⑥"));
        }
    }

    @ParameterizedTest @ValueSource(booleans={false,true})
    void exactPhaseIgnoresLineWrapsSpacesAndStepMetadata(boolean extended) throws Exception {
        try (var recent = new XSSFWorkbook(); var old = new XSSFWorkbook()) {
            recipe(recent, extended, "new", "숙주나물무침", "normalized");
            recipe(old, extended, "old", "STEP3 숙주나물무침\r\n⑤⑥", "wrapped exact");
            assertEquals("wrapped exact", find(extended, recent, old, "숙주나물무침⑤⑥"));
        }
    }

    @ParameterizedTest @ValueSource(booleans={false,true})
    void doesNotReplaceCompoundWithStandaloneSauce(boolean extended) throws Exception {
        try (var recent = new XSSFWorkbook()) {
            recipe(recent, extended, "new", "허니머스타드소스①⑤", "sauce only");
            assertNull(find(extended, recent, null, "훈제오리구이①②&허니머스타드소스①⑤"));
            assertNull(find(extended, recent, null, ""));
        }
    }
}
