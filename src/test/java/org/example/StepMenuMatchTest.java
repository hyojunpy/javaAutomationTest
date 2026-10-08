package org.example;

import java.nio.file.Path;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class StepMenuMatchTest {
    private <T> T renderer(Class<T> type) throws Exception {
        var ctor = type.getDeclaredConstructor(Path.class); ctor.setAccessible(true);
        return ctor.newInstance(Path.of("."));
    }
    @ParameterizedTest
    @ValueSource(strings={"step1 우엉당근죽", "STEP1우엉당근죽", "Step 2\n우엉당근죽", "step3우엉당근죽", "\u00a0step1\u00a0우엉당근죽", "step2\r\n우엉당근죽"})
    void matchesLabelledTemplateAndKeepsOriginalDisplayAndIngredients(String labelled) throws Exception {
        try (var book = new XSSFWorkbook()) {
            var sheet = book.createSheet();
            for (int r=0;r<2;r++) {
                var row=sheet.createRow(r);
                if(r==0) {row.createCell(1).setCellValue(labelled);row.createCell(2).setCellValue(labelled);}
                row.createCell(3).setCellValue(r==0?"우엉":"당근"); row.createCell(5).setCellValue(r==0?15:10);
            }
            sheet.addMergedRegion(new CellRangeAddress(0,1,1,1));
            sheet.addMergedRegion(new CellRangeAddress(0,1,2,2));
            var general=renderer(JsonToExcelGeneral.class);
            var block=general.findBlockInTemplateExact(book,"우엉당근죽",general.colsOf(JsonToExcelGeneral.TplKind.P35)).orElseThrow();
            assertEquals(labelled.trim(),block.displayName);
            assertEquals(java.util.List.of("우엉","당근"),block.items.stream().map(i->i.ingredient).toList());
            var extended=renderer(JsonToExcel.class);
            assertTrue(extended.findBlockInTemplateExactOnly(book,"우엉당근죽").found);
            assertEquals(extended.normalizeForMatch("우엉당근죽"),extended.normalizeForMatch(labelled));
        }
    }
    @ParameterizedTest
    @ValueSource(strings={"step4 우엉당근죽","step10 우엉당근죽","step21 우엉당근죽","우엉step1당근죽","firststep1 우엉당근죽"})
    void doesNotStripOtherNumbersOrMiddleOfName(String name) throws Exception {
        assertEquals(name,MenuNamePrefix.withoutStep(name));
        assertNotEquals(JsonToExcelGeneral.normalizeForMatch("우엉당근죽"),JsonToExcelGeneral.normalizeForMatch(name));
        var extended=renderer(JsonToExcel.class);
        assertNotEquals(extended.normalizeForMatch("우엉당근죽"),extended.normalizeForMatch(name));
    }
    @Test void nullAndOrdinaryNamesRemainSafe() {
        assertEquals("",MenuNamePrefix.withoutStep(null));
        assertEquals("쇠고기우엉주먹밥⑤⑯",MenuNamePrefix.withoutStep("쇠고기우엉주먹밥⑤⑯"));
    }
}
