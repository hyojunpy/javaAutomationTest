package org.example;

import java.nio.file.*;
import java.util.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;
import static org.junit.jupiter.api.Assertions.*;

class ExtendedRecipePolicyTest {
    @TempDir Path home;
    private final HwpToJson parser = new HwpToJson();
    private static final String DUCK = "훈제오리구이①②⑤⑥⑫⑯";
    private static final String SAUCE = "&허니머스타드소스①⑤";

    @ParameterizedTest
    @ValueSource(strings={"훈제오리구이①②⑤⑥⑫⑯\n&허니머스타드소스①⑤",
            "훈제오리구이\n①②⑤⑥⑫⑯\n&허니머스타드소스①⑤",
            "훈제오리구이\n①②⑤⑥⑫⑯&\n허니머스타드소스①⑤"})
    void joinsWrappedCompoundInsideOneCell(String raw) {
        assertEquals(List.of(DUCK + SAUCE, "숙주나물무침⑤⑥"),
                parser.splitMenusByLineThenSlash(raw + "\n숙주나물무침⑤⑥"));
    }

    @Test void doesNotJoinDifferentCellsOrAcrossSlash() {
        assertEquals(List.of(DUCK), parser.splitMenusByLineThenSlash(DUCK));
        assertEquals(List.of(SAUCE), parser.splitMenusByLineThenSlash(SAUCE));
        assertEquals(List.of(DUCK, SAUCE), parser.splitMenusByLineThenSlash(DUCK + "/" + SAUCE));
        assertEquals(List.of(DUCK, "①⑤" + SAUCE), parser.splitMenusByLineThenSlash(DUCK + "/①⑤\n" + SAUCE));
    }

    @ParameterizedTest
    @ValueSource(booleans={true,false})
    void exportsTemplateNamesAndCombinedIngredientsForBothAges(boolean age12) throws Exception {
        Path input = Files.createDirectories(home.resolve("input"));
        Path template = input.resolve(TemplateCatalog.filename(true, age12, true));
        try (var book = new XSSFWorkbook(); var stream = Files.newOutputStream(template)) {
            var old = book.createSheet("26. 2월 넷째주");
            var oldRow = old.createRow(8); oldRow.createCell(1).setCellValue("숙주나물무침⑤⑥");
            oldRow.createCell(2).setCellValue("간장"); oldRow.createCell(4).setCellValue(1);
            var sheet = book.createSheet("26. 9월 다섯째주");
            for (int r=0;r<7;r++) sheet.createRow(r);
            var bean = sheet.createRow(8); bean.createCell(1).setCellValue("숙주나물무침");
            bean.createCell(2).setCellValue("소금"); bean.createCell(3).setCellValue(0.65); bean.createCell(4).setCellValue(1);
            bean.createCell(7).setCellValue("무친다.");
            var duck = sheet.createRow(10); duck.createCell(1).setCellValue(DUCK + "\n" + SAUCE);
            duck.createCell(2).setCellValue("훈제오리"); duck.createCell(3).setCellValue(26); duck.createCell(4).setCellValue(40);
            duck.createCell(7).setCellValue("소스와 함께 제공한다.");
            var sauce = sheet.createRow(11); sauce.createCell(2).setCellValue("머스타드소스");
            sauce.createCell(3).setCellValue(3.25); sauce.createCell(4).setCellValue(5);
            sheet.addMergedRegion(new CellRangeAddress(10,11,1,1));
            book.write(stream);
        }
        var plan = GoldenFiles.JSON.createArrayNode(); var day = plan.addObject(); day.put("date",25); day.put("weekday","수");
        var menus = day.putArray("menus"); menus.add("숙주나물무침⑤⑥");
        parser.splitMenusByLineThenSlash(DUCK + "\n" + SAUCE).forEach(menus::add);
        menus.add("없는나물⑤");
        Path output = home.resolve("output.xlsx");
        JsonToExcel.convertFromJsonString(plan.toString(),Path.of("2026년 2월 시간연장형("+(age12?"만1-2세":"만3-5세")+").hwp"),output,home);
        try (var book = WorkbookResources.open(output)) {
            List<String> names = new ArrayList<>(), ingredients = new ArrayList<>();
            for (var row : book.getSheetAt(0)) {
                var menu = row.getCell(1); if (menu != null) names.add(menu.toString());
                var item = row.getCell(2); if (item != null) ingredients.add(item.toString());
            }
            assertTrue(names.contains("숙주나물무침"));
            assertFalse(names.contains("숙주나물무침⑤⑥"));
            assertTrue(names.contains(DUCK + "\n" + SAUCE));
            assertFalse(names.contains(SAUCE));
            assertTrue(names.contains("없는나물⑤"));
            assertTrue(ingredients.containsAll(List.of("소금","훈제오리","머스타드소스")));
            assertFalse(ingredients.contains("간장"));
        }
    }
}
