// (full file) src/main/java/org/example/JsonToExcel.java
package org.example;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.*;
import java.time.YearMonth;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * JSON(menu plan) + XLSX(template) -> XLSX(output)
 *
 * Single sheet output:
 *  - Append day blocks vertically with 4 blank rows between blocks.
 *  - Global parameters on D2 (multiplier for 1~2) and D3 (multiplier for 3~5).
 *
 * Template columns (both 3-5 and 1-2 share base):
 *   B: Menu, C: Ingredient, D: 1-serving(1~2), E: 1-serving(3~5),
 *   F: Total(1~2), G: Total(3~5), H..L: Method
 */
public class JsonToExcel {

    // ===== Input JSON (fallback only) =====
    static final Path JSON_PLAN = Paths.get("output/25.12. 만1-2세 시간연장형.json");


    static final String TEMPLATE_35_NAME = "★2021.9~ 조리지시서(만3-5세 시간연장형).xlsx";
    static final String TEMPLATE_12_NAME = "★2021.9~ 조리지시서(만1-2세 시간연장형).xlsx";

    static final String[] HEADER_35 = {
            "메뉴명","식재료명","1인 제공량(g)\n1~2세","1인 제공량(g)\n3~5세",
            "총 발주량\n1~2세","총 발주량\n3~5세","만드는방법"
    };
    static final String[] HEADER_12 = {
            "메뉴명","식재료명","1인 제공량(g)","총 발주량","만드는방법"
    };

    static final int COL_MENU   = 1;
    static final int COL_ING    = 2;
    static final int COL_P12    = 3;
    static final int COL_P35    = 4;
    static final int COL_T12    = 5;
    static final int COL_T35    = 6;
    static final int COL_METHOD = 7;

    static final int GAP_ROWS = 4;

    enum OutputMode { AGE12, AGE35 }

    static final Map<String,String> ALIAS = new HashMap<>();
    static {
        ALIAS.put("쇠고기", "소고기");
        ALIAS.put("계란", "달걀");
        ALIAS.put("닭 살", "닭살");
    }

    public static void main(String[] args) throws Exception {
        ZipSecureFile.setMinInflateRatio(0.0d);
        ZipSecureFile.setMaxFileCount(20000);

        Path jsonPlan;
        if (args != null && args.length >= 1 && args[0] != null && args[0].trim().length() > 0) {
            jsonPlan = Paths.get(args[0]).toAbsolutePath();
        } else {
            jsonPlan = JSON_PLAN.toAbsolutePath();
        }

        Path outPath;
        if (args != null && args.length >= 2 && args[1] != null && args[1].trim().length() > 0) {
            outPath = Paths.get(args[1]).toAbsolutePath();
        } else {
            outPath = deriveOutXlsxPathFromJson(jsonPlan, Paths.get("output")).toAbsolutePath();
        }

        if (outPath.getParent() != null) Files.createDirectories(outPath.getParent());

        YearMonth ym = inferYMFromJson(jsonPlan);
        String titlePrefix = String.format("%d년 %d월", ym.getYear(), ym.getMonthValue());

        OutputMode mode = inferModeFromJsonName(jsonPlan);
        Path chosenTemplate = openTemplateForMode(mode);

        List<DayPlan> plan = readPlan(jsonPlan);

        try (Workbook template = WorkbookFactory.create(Files.newInputStream(chosenTemplate));
             Workbook out = new XSSFWorkbook()) {

            Styles S = Styles.build(out);

            Sheet sh = out.createSheet("전체");
            int nextRow = 0;

            for (DayPlan d : plan) {
                if (d == null) continue;
                nextRow = writeOneDay(sh, S, d, template, titlePrefix, mode, nextRow);
                nextRow += GAP_ROWS;
            }

            int[] widths = { 5000, 5000, 3500, 3500, 3500, 3500, 4500, 4500, 4500, 4500, 4500 };
            for (int i = 0; i < widths.length; i++) sh.setColumnWidth(i, widths[i]);

            try (OutputStream os = Files.newOutputStream(outPath)) {
                out.write(os);
            }
        }

        System.out.println("DONE → " + outPath.toAbsolutePath());
    }

    // English comment: Resolve template from installed app location (app.home/input) first
    static Path openTemplateForMode(OutputMode mode) throws Exception {
        Path appHome = getAppHomeDir();
        Path inputDir = appHome.resolve("input");

        Path candidate;
        if (mode == OutputMode.AGE12) candidate = inputDir.resolve(TEMPLATE_12_NAME);
        else candidate = inputDir.resolve(TEMPLATE_35_NAME);

        if (Files.exists(candidate)) return candidate;

        // Fallback: current working directory ./input
        Path fallbackDir = Paths.get("").toAbsolutePath().resolve("input");
        if (mode == OutputMode.AGE12) candidate = fallbackDir.resolve(TEMPLATE_12_NAME);
        else candidate = fallbackDir.resolve(TEMPLATE_35_NAME);

        if (Files.exists(candidate)) return candidate;

        throw new IllegalStateException("Template not found: " + candidate.toAbsolutePath());
    }

    // English comment: Get app home directory for packaged app
    static Path getAppHomeDir() {
        String home = System.getProperty("app.home");
        if (home != null && home.trim().length() > 0) {
            return Paths.get(home).toAbsolutePath();
        }
        return Paths.get("").toAbsolutePath();
    }

    static OutputMode inferModeFromJsonName(Path jsonPath) {
        String name = jsonPath.getFileName().toString().toLowerCase(Locale.ROOT);
        boolean is12 = name.contains("만1-2세") || name.contains("만1~2세") || name.contains("만1–2세");
        if (is12) return OutputMode.AGE12;
        return OutputMode.AGE35;
    }

    static int writeOneDay(Sheet sh, Styles S, DayPlan d, Workbook template,
                           String titlePrefix, OutputMode mode, int startRow) {
        int r = startRow;
        int methodFirst;
        if (mode == OutputMode.AGE35) methodFirst = 6;
        else methodFirst = 4;

        int methodLast  = methodFirst + 4;
        int lastColForBlock = methodLast;

        // Title
        Row tr = safeRow(sh, r++);
        int headerRowIndex = tr.getRowNum();
        tr.setHeightInPoints(22);
        Cell t0 = tr.createCell(0);
        t0.setCellValue(String.format("%s %d일 (%s)", titlePrefix, d.date, d.weekday));
        t0.setCellStyle(S.title);
        CellRangeAddress titleMerge = new CellRangeAddress(tr.getRowNum(), tr.getRowNum(), 0, lastColForBlock);
        sh.addMergedRegion(titleMerge);
        setMergedBorder(sh, titleMerge, BorderStyle.THIN);

        // Params
        Row pr = safeRow(sh, r++);
        Cell d2 = pr.createCell(3);
        d2.setCellStyle(S.param);
        d2.setCellValue(1.0);

        Row pr2 = safeRow(sh, r++);
        Cell d3 = pr2.createCell(3);
        d3.setCellStyle(S.param);
        d3.setCellValue(1.0);

        // Header
        Row hr = safeRow(sh, r++);
        if (mode == OutputMode.AGE35) {
            for (int i = 0; i < HEADER_35.length; i++) {
                Cell hc = hr.createCell(i);
                hc.setCellValue(HEADER_35[i]);
                hc.setCellStyle(S.header);
            }
        } else {
            for (int i = 0; i < HEADER_12.length; i++) {
                Cell hc = hr.createCell(i);
                hc.setCellValue(HEADER_12[i]);
                hc.setCellStyle(S.header);
            }
        }

        for (int c = methodFirst; c <= methodLast; c++) {
            Cell hc = hr.getCell(c);
            if (hc == null) hc = hr.createCell(c);
            hc.setCellStyle(S.header);
        }
        CellRangeAddress headerMethod = new CellRangeAddress(hr.getRowNum(), hr.getRowNum(), methodFirst, methodLast);
        sh.addMergedRegion(headerMethod);
        setMergedBorder(sh, headerMethod, BorderStyle.THIN);
        hr.setHeightInPoints(22);

        // Menus
        List<String> head = new ArrayList<>();
        List<String> tail = new ArrayList<>();
        if (d.menus != null) {
            for (String m : d.menus) {
                if (isTailMenu(m)) tail.add(m);
                else head.add(m);
            }
        }
        List<String> orderedMenus = new ArrayList<>(head);
        orderedMenus.addAll(tail);

        for (String rawMenu : orderedMenus) {
            SearchResult sr = findBlockInTemplateExactOnly(template, rawMenu);
            int blockStart = r;

            if (!sr.found) {
                Row row = safeRow(sh, r++);
                writeDataRowWithFormulas(row, S, rawMenu, "", "", "", "", "", "", mode);
            } else {
                Block b = sr.block;

                String menuTitleToShow;
                if (sr.matchedByExact) menuTitleToShow = nz(b.displayName);
                else menuTitleToShow = rawMenu;

                boolean noItemsAndNoMethod = b.items.isEmpty() && isBlank(b.method);
                if (noItemsAndNoMethod) {
                    Row row = safeRow(sh, r++);
                    writeDataRowWithFormulas(row, S, menuTitleToShow, "", "", "", "", "", "", mode);
                } else {
                    for (int i = 0; i < b.items.size(); i++) {
                        Item it = b.items.get(i);
                        Row row = safeRow(sh, r++);
                        String menuCell = "";
                        String method = "";
                        if (i == 0) {
                            menuCell = menuTitleToShow;
                            method = compactSpaces(b.method);
                        }
                        writeDataRowWithFormulas(row, S,
                                menuCell, it.ingredient, it.p12, it.p35, it.t12, it.t35, method, mode);
                    }
                }
            }

            int blockEnd = r - 1;

            // Merge menu column A
            Cell aTop = safeCell(sh, blockStart, 0);
            aTop.setCellStyle(S.bodyCenter);
            if (blockEnd > blockStart) {
                CellRangeAddress aMerge = new CellRangeAddress(blockStart, blockEnd, 0, 0);
                sh.addMergedRegion(aMerge);
                setMergedBorder(sh, aMerge, BorderStyle.THIN);
            } else {
                setBorders(aTop, BorderStyle.THIN);
            }

            // Merge method area and apply styles
            CellRangeAddress methodMerge = new CellRangeAddress(blockStart, blockEnd, methodFirst, methodLast);
            sh.addMergedRegion(methodMerge);
            setMergedBorder(sh, methodMerge, BorderStyle.THIN);
            for (int rr = blockStart; rr <= blockEnd; rr++) {
                for (int cc = methodFirst; cc <= methodLast; cc++) {
                    Cell c = safeCell(sh, rr, cc);
                    c.setCellStyle(S.methodMerged);
                }
            }

            if (mode == OutputMode.AGE35) {
                for (int br = blockStart; br <= blockEnd; br++) {
                    for (int bc = 1; bc <= 5; bc++) setBorders(safeCell(sh, br, bc), BorderStyle.THIN);
                    setBorders(safeCell(sh, br, 0), BorderStyle.THIN);
                }
            } else {
                for (int br = blockStart; br <= blockEnd; br++) {
                    for (int bc = 1; bc <= 3; bc++) setBorders(safeCell(sh, br, bc), BorderStyle.THIN);
                    setBorders(safeCell(sh, br, 0), BorderStyle.THIN);
                }
            }
        }

        int lastDataRow = r - 1;
        if (lastDataRow >= headerRowIndex) {
            CellRangeAddress outer = new CellRangeAddress(headerRowIndex, lastDataRow, 0, lastColForBlock);
            setMergedBorder(sh, outer, BorderStyle.DOUBLE);
        }

        return r;
    }

    static void writeDataRowWithFormulas(Row row, Styles S,
                                         String menu, String ing,
                                         String p12, String p35, String t12, String t35,
                                         String method, OutputMode mode) {
        int excelRow = row.getRowNum() + 1;
        row.setHeightInPoints(18);

        Cell a = row.createCell(0);
        a.setCellValue(nz(menu));
        a.setCellStyle(S.bodyCenter);

        Cell b = row.createCell(1);
        b.setCellValue(nz(ing));
        b.setCellStyle(S.bodyCenter);

        if (mode == OutputMode.AGE35) {
            Cell d = row.createCell(3);
            boolean dIsNum = trySetNumeric(d, p35);
            if (!dIsNum) d.setCellValue(nz(p35));
            d.setCellStyle(S.numGeneral);

            Cell c = row.createCell(2);
            c.setCellFormula(String.format("D%d*0.65", excelRow));

            boolean forceOneDecimal = shouldForceOneDecimal(p35);
            if (forceOneDecimal) c.setCellStyle(S.num1dec);
            else c.setCellStyle(S.numGeneral);

            Cell e = row.createCell(4);
            e.setCellFormula(String.format("C%d*$D$2", excelRow));
            e.setCellStyle(S.numGeneral);

            Cell f = row.createCell(5);
            f.setCellFormula(String.format("D%d*$D$3", excelRow));
            f.setCellStyle(S.numGeneral);

            Cell g = row.createCell(6);
            g.setCellValue(compactSpaces(nz(method)));
        } else {
            Cell c = row.createCell(2);
            boolean cNum = trySetNumeric(c, p12);
            if (!cNum) {
                Double p35val = parseNumericOrNull(p35);
                if (p35val != null) {
                    double v = p35val * 0.65;
                    c.setCellValue(v);
                    boolean forceOneDecimal = shouldForceOneDecimal(p35);
                    if (forceOneDecimal) c.setCellStyle(S.num1dec);
                    else c.setCellStyle(S.numGeneral);
                } else {
                    c.setCellValue(nz(p12));
                    c.setCellStyle(S.num1dec);
                }
            } else {
                c.setCellStyle(S.num1dec);
            }

            Cell d = row.createCell(3);
            d.setCellFormula(String.format("C%d*$D$2", excelRow));
            d.setCellStyle(S.numGeneral);

            Cell e = row.createCell(4);
            e.setCellValue(compactSpaces(nz(method)));
        }
    }

    static boolean shouldForceOneDecimal(String p35) {
        Double v = parseNumericOrNull(p35);
        if (v == null) return true;
        double x = v * 0.65;
        double xTimes10 = Math.round(x * 10.0) / 10.0;
        return Math.abs(xTimes10 - Math.rint(xTimes10)) > 1e-9;
    }

    static class SearchResult {
        final boolean found;
        final boolean matchedByExact;
        final Block block;
        SearchResult(boolean f, boolean exact, Block b){ this.found = f; this.matchedByExact = exact; this.block = b; }
        static SearchResult miss(){ return new SearchResult(false, false, null); }
    }

    static SearchResult findBlockInTemplateExactOnly(Workbook wb, String menuRaw){
        String keyExact = compactSpacesPreserveAll(applyAlias(menuRaw));
        String keyNorm  = normalizeForMatch(keyExact);

        for (int s = wb.getNumberOfSheets() - 1; s >= 0; s--) {
            Sheet sh = wb.getSheetAt(s);
            int lastRow = sh.getLastRowNum();

            int r = 0;
            while (r <= lastRow) {
                String cellMenu = readStringConsideringMerged(sh, r, COL_MENU);
                if (isBlank(cellMenu)) { r++; continue; }

                CellRangeAddress menuRange = findMergedRange(sh, r, COL_MENU);
                int first = r;
                int last = r;
                if (menuRange != null) {
                    first = menuRange.getFirstRow();
                    last  = menuRange.getLastRow();
                } else {
                    int rr = r + 1;
                    while (rr <= lastRow) {
                        String nextMenu = readStringConsideringMerged(sh, rr, COL_MENU);
                        if (!isBlank(nextMenu)) break;
                        rr++;
                    }
                    last = rr - 1;
                }

                String cellExact = compactSpacesPreserveAll(cellMenu);
                String cellNorm  = normalizeForMatch(cellExact);

                if (cellExact.equals(keyExact)) return new SearchResult(true, true, readBlock(sh, first, last));
                if (cellNorm.equals(keyNorm))  return new SearchResult(true, false, readBlock(sh, first, last));

                r = last + 1;
            }
        }
        return SearchResult.miss();
    }

    static Block readBlock(Sheet sh, int first, int last){
        String methodTop = "";
        outer:
        for (int rr = first; rr <= last; rr++) {
            for (int cc = COL_METHOD; cc <= COL_METHOD + 4; cc++) {
                String v = readStringConsideringMerged(sh, rr, cc);
                if (!isBlank(v)) { methodTop = v.trim(); break outer; }
            }
        }
        List<Item> items = new ArrayList<>();
        for (int rr = first; rr <= last; rr++) {
            String ing = readStringConsideringMerged(sh, rr, COL_ING);
            String c12 = readStringConsideringMerged(sh, rr, COL_P12);
            String c35 = readStringConsideringMerged(sh, rr, COL_P35);
            String t12 = readStringConsideringMerged(sh, rr, COL_T12);
            String t35 = readStringConsideringMerged(sh, rr, COL_T35);
            boolean rowEmpty = isBlank(ing) && isBlank(c12) && isBlank(c35) && isBlank(t12) && isBlank(t35);
            if (!rowEmpty && !"식재료명".equals(ing)) {
                Item it = new Item();
                it.ingredient = ing;
                it.p12 = c12;
                it.p35 = c35;
                it.t12 = t12;
                it.t35 = t35;
                items.add(it);
            }
        }
        Block b = new Block();
        b.displayName = readStringConsideringMerged(sh, first, COL_MENU);
        b.method = nz(methodTop);
        b.items  = items;
        return b;
    }

    static Path deriveOutXlsxPathFromJson(Path jsonPath, Path outDir){
        String file = jsonPath.getFileName().toString();
        int dot = file.lastIndexOf('.');
        String stem = (dot > 0) ? file.substring(0, dot) : file;
        String outBase = stem + "_수정";
        return outDir.resolve(outBase + ".xlsx");
    }

    static CellRangeAddress findMergedRange(Sheet sh, int r, int c){
        for (int i = 0; i < sh.getNumMergedRegions(); i++) {
            CellRangeAddress ra = sh.getMergedRegion(i);
            if (ra.isInRange(r, c)) return ra;
        }
        return null;
    }

    static String readStringConsideringMerged(Sheet sh, int r, int c){
        Cell cell = getMergedAnchorCell(sh, r, c);
        if (cell == null) return "";
        return getString(cell).trim();
    }

    static Cell getMergedAnchorCell(Sheet sh, int r, int c){
        for (int i=0;i<sh.getNumMergedRegions();i++){
            CellRangeAddress ra = sh.getMergedRegion(i);
            if (ra.isInRange(r,c)){
                Row topRow = sh.getRow(ra.getFirstRow());
                if (topRow == null) return null;
                return topRow.getCell(ra.getFirstColumn());
            }
        }
        Row row = sh.getRow(r);
        if (row == null) return null;
        return row.getCell(c);
    }

    static String getString(Cell cell){
        if (cell == null) return "";
        switch (cell.getCellType()){
            case STRING:  return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) return cell.getDateCellValue().toString();
                double v = cell.getNumericCellValue();
                long rnd = Math.round(v);
                if (Math.abs(v - rnd) < 1e-9) return String.valueOf(rnd);
                return String.valueOf(v);
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try { return cell.getStringCellValue(); }
                catch (Exception e) {
                    try { return String.valueOf(cell.getNumericCellValue()); }
                    catch (Exception ignore) { return ""; }
                }
            default: return "";
        }
    }

    static boolean trySetNumeric(Cell c, String s) {
        if (isBlank(s)) return false;
        try {
            String t = s
                    .replace(",", "")
                    .replace("\u00A0", "")
                    .replaceAll("[\\u2000-\\u200B\\u202F\\u205F\\u3000]", "")
                    .replace('．','.')
                    .replace('。','.')
                    .replace('･','.')
                    .replace('·','.')
                    .trim();
            t = t.replaceAll("[^0-9.\\-]", "");
            t = t.replaceAll("\\.(?=\\.)", "");
            t = t.replaceAll("\\.$", "");
            if (!t.matches("[-]?(\\d+\\.?\\d*|\\d*\\.\\d+)")) return false;

            double v = Double.parseDouble(t);
            c.setCellValue(v);
            return true;
        } catch (Exception e) { return false; }
    }

    static Double parseNumericOrNull(String s) {
        if (isBlank(s)) return null;
        try {
            String t = s
                    .replace(",", "")
                    .replace("\u00A0", "")
                    .replaceAll("[\\u2000-\\u200B\\u202F\\u205F\\u3000]", "")
                    .replace('．','.')
                    .replace('。','.')
                    .replace('･','.')
                    .replace('·','.')
                    .trim();
            t = t.replaceAll("[^0-9.\\-]", "");
            t = t.replaceAll("\\.(?=\\.)", "");
            t = t.replaceAll("\\.$", "");
            if (!t.matches("[-]?(\\d+\\.?\\d*|\\d*\\.\\d+)")) return null;
            return Double.parseDouble(t);
        } catch (Exception e) {
            return null;
        }
    }

    static void setBorders(Cell cell, BorderStyle bs) {
        CellStyle clone = cell.getSheet().getWorkbook().createCellStyle();
        clone.cloneStyleFrom(cell.getCellStyle());
        clone.setBorderBottom(bs);
        clone.setBorderTop(bs);
        clone.setBorderLeft(bs);
        clone.setBorderRight(bs);
        cell.setCellStyle(clone);
    }

    static void setMergedBorder(Sheet sh, CellRangeAddress rgn, BorderStyle bs) {
        org.apache.poi.ss.util.RegionUtil.setBorderTop(bs, rgn, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderBottom(bs, rgn, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderLeft(bs, rgn, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderRight(bs, rgn, sh);
    }

    static Row safeRow(Sheet sh, int r){
        Row row = sh.getRow(r);
        if (row == null) row = sh.createRow(r);
        return row;
    }

    static Cell safeCell(Sheet sh, int r, int c) {
        Row row = safeRow(sh, r);
        Cell cell = row.getCell(c);
        if (cell == null) cell = row.createCell(c);
        return cell;
    }

    static boolean isBlank(String s){ return s == null || s.trim().isEmpty(); }
    static String nz(String s){ return s == null ? "" : s; }

    static String compactSpaces(String s){
        if (s == null) return "";
        return s.replace('\u00A0',' ')
                .replaceAll("[ \\t]{2,}", " ")
                .trim();
    }

    static String compactSpacesPreserveAll(String s){
        if (s == null) return "";
        return s.replace('\u00A0',' ')
                .replaceAll("[ \\t]{2,}", " ")
                .trim();
    }

    static String applyAlias(String s){
        if (s == null) return "";
        String t = s;
        for (Map.Entry<String,String> e : ALIAS.entrySet()) {
            t = t.replace(e.getKey(), e.getValue());
        }
        return t;
    }

    static String normalizeForMatch(String s){
        if (s == null) return "";
        String t = applyAlias(s);
        t = t.replaceAll("[\u2460-\u2473①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲]", "");
        t = t.replaceAll("[()\\[\\]{}／/·ㆍ・＆&\\-★☆※•….,;:]", "");
        t = t.replaceAll("\\s+", "");
        return t;
    }

    static YearMonth inferYMFromJson(Path jsonPath) {
        String name = jsonPath.getFileName().toString();
        Matcher mYY = Pattern.compile("^(\\d{2})\\.(\\d{1,2})").matcher(name);
        if (mYY.find()) {
            int yy = Integer.parseInt(mYY.group(1));
            int mm = Integer.parseInt(mYY.group(2));
            return YearMonth.of(2000 + yy, Math.min(12, Math.max(1, mm)));
        }
        Matcher mYYYY = Pattern.compile("(20\\d{2})[\\.-_/ ]?(\\d{1,2})").matcher(name);
        if (mYYYY.find()) {
            int y = Integer.parseInt(mYYYY.group(1));
            int mo = Integer.parseInt(mYYYY.group(2));
            return YearMonth.of(y, Math.min(12, Math.max(1, mo)));
        }
        return YearMonth.of(2025, 12);
    }

    static boolean isTailMenu(String s){
        if (s == null) return false;
        String t = s.replaceAll("[\u2460-\u2473①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲]", "");
        t = t.replaceAll("\\s+", "");
        return t.contains("배추김치") || t.contains("깍두기");
    }

    public static class DayPlan {
        public int date;
        public String weekday;
        public List<String> menus;
    }
    static class Item { String ingredient, p12, p35, t12, t35; }
    static class Block { String displayName; String method; List<Item> items = new ArrayList<>(); }

    static class Styles {
        final CellStyle title;
        final CellStyle header;
        final CellStyle bodyCenter;
        final CellStyle methodMerged;
        final CellStyle num1dec;
        final CellStyle numGeneral;
        final CellStyle param;

        Styles(CellStyle title, CellStyle header, CellStyle bodyCenter,
               CellStyle methodMerged, CellStyle num1dec, CellStyle numGeneral, CellStyle param) {
            this.title = title;
            this.header = header;
            this.bodyCenter = bodyCenter;
            this.methodMerged = methodMerged;
            this.num1dec = num1dec;
            this.numGeneral = numGeneral;
            this.param = param;
        }

        static Styles build(Workbook wb) {
            DataFormat df = wb.createDataFormat();

            Font bodyFont = wb.createFont();
            bodyFont.setFontName("한컴산뜻돋움");
            bodyFont.setFontHeightInPoints((short) 10);

            Font headerFont = wb.createFont();
            headerFont.setFontName("한컴산뜻돋움");
            headerFont.setFontHeightInPoints((short) 12);
            headerFont.setBold(true);

            CellStyle base = wb.createCellStyle();
            base.setBorderBottom(BorderStyle.THIN);
            base.setBorderTop(BorderStyle.THIN);
            base.setBorderLeft(BorderStyle.THIN);
            base.setBorderRight(BorderStyle.THIN);

            CellStyle title = wb.createCellStyle();
            title.cloneStyleFrom(base);
            title.setAlignment(HorizontalAlignment.CENTER);
            title.setVerticalAlignment(VerticalAlignment.CENTER);
            title.setWrapText(true);
            title.setFont(headerFont);

            CellStyle header = wb.createCellStyle();
            header.cloneStyleFrom(base);
            header.setAlignment(HorizontalAlignment.CENTER);
            header.setVerticalAlignment(VerticalAlignment.CENTER);
            header.setWrapText(true);
            header.setFont(headerFont);

            CellStyle bodyCenter = wb.createCellStyle();
            bodyCenter.cloneStyleFrom(base);
            bodyCenter.setAlignment(HorizontalAlignment.CENTER);
            bodyCenter.setVerticalAlignment(VerticalAlignment.CENTER);
            bodyCenter.setWrapText(true);
            bodyCenter.setFont(bodyFont);

            CellStyle methodMerged = wb.createCellStyle();
            methodMerged.cloneStyleFrom(base);
            methodMerged.setAlignment(HorizontalAlignment.LEFT);
            methodMerged.setVerticalAlignment(VerticalAlignment.CENTER);
            methodMerged.setWrapText(true);
            methodMerged.setFont(bodyFont);

            CellStyle num1dec = wb.createCellStyle();
            num1dec.cloneStyleFrom(bodyCenter);
            num1dec.setDataFormat(df.getFormat("0.0"));
            num1dec.setFont(bodyFont);

            CellStyle numGeneral = wb.createCellStyle();
            numGeneral.cloneStyleFrom(bodyCenter);
            numGeneral.setDataFormat(df.getFormat("General"));
            numGeneral.setFont(bodyFont);

            CellStyle param = wb.createCellStyle();
            param.cloneStyleFrom(bodyCenter);
            param.setDataFormat(df.getFormat("General"));
            param.setFont(bodyFont);

            return new Styles(title, header, bodyCenter, methodMerged, num1dec, numGeneral, param);
        }
    }

    static List<DayPlan> readPlan(Path json) throws IOException {
        ObjectMapper om = new ObjectMapper();
        try (var is = Files.newInputStream(json)) {
            return om.readValue(is, new TypeReference<List<DayPlan>>() {});
        }
    }

    // New: JSON -> Excel with explicit paths
    public static Path convert(Path jsonPath, Path outXlsxPath) throws Exception {
        ZipSecureFile.setMinInflateRatio(0.0d);
        ZipSecureFile.setMaxFileCount(20000);

        YearMonth ym = inferYMFromJson(jsonPath);
        String titlePrefix = String.format("%d년 %d월", ym.getYear(), ym.getMonthValue());

        if (outXlsxPath.getParent() != null) Files.createDirectories(outXlsxPath.getParent());

        OutputMode mode = inferModeFromJsonName(jsonPath);
        Path chosenTemplate = openTemplateForMode(mode);

        List<DayPlan> plan = readPlan(jsonPath);

        try (Workbook template = WorkbookFactory.create(Files.newInputStream(chosenTemplate));
             Workbook out = new XSSFWorkbook()) {

            Styles S = Styles.build(out);
            Sheet sh = out.createSheet("전체");

            int nextRow = 0;
            for (DayPlan d : plan) {
                if (d == null) continue;
                nextRow = writeOneDay(sh, S, d, template, titlePrefix, mode, nextRow);
                nextRow += GAP_ROWS;
            }

            int[] widths = { 5000, 5000, 3500, 3500, 3500, 3500, 4500, 4500, 4500, 4500, 4500 };
            for (int i = 0; i < widths.length; i++) sh.setColumnWidth(i, widths[i]);

            try (OutputStream os = Files.newOutputStream(outXlsxPath)) {
                out.write(os);
            }
        }

        return outXlsxPath;
    }
}
