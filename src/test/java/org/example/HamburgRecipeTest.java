package org.example;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import java.nio.file.*;
import java.util.*;
import static org.junit.jupiter.api.Assertions.*;

class HamburgRecipeTest {
    @TempDir Path home;

    @Test void correctedMenuLoadsThreeIngredientBlockWithAmountsAndFormulas() throws Exception {
        Path input = Files.createDirectories(home.resolve("input"));
        Path template = input.resolve(TemplateCatalog.filename(false, false, true));
        String name = "함박스테이크①②⑤⑥⑩⑫⑮⑯⑱&소스①⑤⑥⑩⑫⑮⑯";
        List<String> ingredients = List.of("함박스테이크", "스테이크소스", "콩기름");
        List<Double> amounts = List.of(45.0, 10.0, 2.0);
        try (var book = new XSSFWorkbook(); var stream = Files.newOutputStream(template)) {
            var sheet = book.createSheet("26. 1월 셋째주");
            for (int r = 0; r < 6; r++) sheet.createRow(r);
            sheet.getRow(1).createCell(4).setCellValue(10);
            sheet.getRow(2).createCell(4).setCellValue(20);
            for (int i = 0; i < ingredients.size(); i++) {
                var row = sheet.createRow(78 + i);
                if (i == 0) {
                    row.createCell(2).setCellValue(name);
                    row.createCell(8).setCellValue("① 함박스테이크를 굽는다. ② 소스를 부어 제공한다.");
                }
                row.createCell(3).setCellValue(ingredients.get(i));
                row.createCell(4).setCellFormula("F" + (79 + i) + "*0.65");
                row.createCell(5).setCellValue(amounts.get(i));
                row.createCell(6).setCellFormula("E" + (79 + i) + "*$E$2");
                row.createCell(7).setCellFormula("F" + (79 + i) + "*$E$3");
            }
            sheet.addMergedRegion(CellRangeAddress.valueOf("C79:C81"));
            sheet.addMergedRegion(CellRangeAddress.valueOf("I79:M81"));
            book.write(stream);
        }
        var menus = new HwpToJsonGeneral().splitMenusByLineThenSlash(
                "함박스테이크\n①②⑤⑥⑩⑫⑮⑯⑱&소스①⑤⑥⑩⑫⑮⑯");
        assertEquals(List.of(name), menus);
        var plan = GoldenFiles.JSON.createArrayNode();
        var day = plan.addObject(); day.put("date", 14); day.put("weekday", "수"); day.put("morningCount", 1);
        var mealList = day.putArray("menus"); mealList.add("파인애플"); menus.forEach(mealList::add);
        day.putArray("pmDesert");
        Path output = home.resolve("result.xlsx");
        JsonToExcelGeneral.convertFromJsonString(plan.toString(), Path.of("2026년 1월 일반형(만3-5세).hwp"), output, home);
        try (var book = WorkbookResources.open(output)) {
            var sheet = book.getSheetAt(0);
            List<String> actualIngredients = new ArrayList<>();
            for (var row : sheet) {
                var cell = row.getCell(3);
                if (cell == null || cell.getCellType() != org.apache.poi.ss.usermodel.CellType.STRING
                        || !ingredients.contains(cell.getStringCellValue())) continue;
                int index = ingredients.indexOf(cell.getStringCellValue());
                actualIngredients.add(cell.getStringCellValue());
                assertEquals(amounts.get(index), row.getCell(5).getNumericCellValue());
                int excelRow = row.getRowNum() + 1;
                assertEquals("F" + excelRow + "*0.65", row.getCell(4).getCellFormula());
                assertEquals("E" + excelRow + "*$E$2", row.getCell(6).getCellFormula());
                assertEquals("F" + excelRow + "*$E$3", row.getCell(7).getCellFormula());
                if (index == 0) assertEquals(name, row.getCell(2).getStringCellValue());
            }
            assertEquals(ingredients, actualIngredients);
        }
    }
}
