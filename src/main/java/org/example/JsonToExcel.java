// (full file) src/main/java/org/example/JsonToExcel.java
package org.example;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.*;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.YearMonth;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * JSON(menu plan) + XLSX(template) -> XLSX(output)
 * <p>
 * Single sheet output:
 * - Append day blocks vertically with 4 blank rows between blocks.
 * - Global parameters on D2 (multiplier for 1~2) and D3 (multiplier for 3~5).
 * <p>
 * Template columns (both 3-5 and 1-2 share base):
 * B: Menu, C: Ingredient, D: 1-serving(1~2), E: 1-serving(3~5),
 * F: Total(1~2), G: Total(3~5), H..L: Method
 */
public class JsonToExcel {
    private final Path appHome;

    private JsonToExcel() { this(TemplateCatalog.defaultHome()); }

    private JsonToExcel(Path appHome) { this.appHome = appHome.toAbsolutePath(); }


    // ===== Input JSON (fallback only) =====
    static final Path JSON_PLAN = Paths.get("output/26.01. 만1-2세 시간연장형.json");



    static final int COL_MENU = 1;
    static final int COL_ING = 2;
    static final int COL_P12 = 3;
    static final int COL_P35 = 4;
    static final int COL_T12 = 5;
    static final int COL_T35 = 6;
    static final int COL_METHOD = 7;

    static final int GAP_ROWS = 2;

    enum OutputMode {AGE12, AGE35}

    final Map<String, CellStyle> STYLE_CACHE = new HashMap<>();

    static final Map<String, String> ALIAS = new HashMap<>();

    static {
        ALIAS.put("쇠고기", "소고기");
        ALIAS.put("계란", "달걀");
        ALIAS.put("닭 살", "닭살");
    }

    private void runMain(String[] args) throws Exception {
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

        doConvert(jsonPlan, outPath);
        System.out.println("DONE → " + outPath.toAbsolutePath());
    }

    // English comment: Get app home directory for packaged app
    Path getAppHomeDir() { return appHome; }

    OutputMode inferModeFromJsonName(Path jsonPath) {
        return SourceMetadata.from(jsonPath).age12() ? OutputMode.AGE12 : OutputMode.AGE35;
    }

    int writeOneDay(Sheet sh, Styles S, DayPlan d, Workbook tplNew, Workbook tplOld, String titlePrefix, OutputMode mode, int startRow) {
        sh.setDefaultColumnStyle(0, S.blankA);

        boolean hasRealMenu = false;
        if (d.menus != null) {
            for (String m : d.menus) {
                String t = nz(m).trim();
                if (t.length() > 0 && !t.equals("없음")) hasRealMenu = true;
            }
        }
        if (!hasRealMenu) return startRow;

        int r = startRow;
        int methodFirst;
        // English comment: Method area should be H~L when output columns are shifted right by 1.
        if (mode == OutputMode.AGE35) methodFirst = 7; // H
        else methodFirst = 5;                          // F (AGE12)
        int methodLast = methodFirst + 4;
        int lastColForBlock = methodLast;


        // Title
        Row tr = safeRow(sh, r++);
        int headerRowIndex = tr.getRowNum();
        tr.setHeightInPoints(21);

        Cell t0 = tr.createCell(1);
        t0.setCellValue(String.format("%s %d일 (%s)", titlePrefix, d.date, d.weekday));
        t0.setCellStyle(S.title);

        CellRangeAddress titleMerge = new CellRangeAddress(tr.getRowNum(), tr.getRowNum(), 1, lastColForBlock);
        sh.addMergedRegion(titleMerge);
        setMergedBorder(sh, titleMerge, BorderStyle.THIN);

        Cell aTitle = safeCell(sh, tr.getRowNum(), 0);
        aTitle.setCellValue("");
        aTitle.setCellStyle(S.blankA);

        // Header
        // Header (2 rows)
// English comment: Column map (0-based):
// A: blank(0), B: menu(1), C: ingredient(2), D: p12(3), E: p35(4),
// F: t12(5), G: t35(6), H~L: method(7~11)

        Row hr1 = safeRow(sh, r++);
        Row hr2 = safeRow(sh, r++);
        removeHorizontalMergesInColumnA(sh, hr1.getRowNum(), hr2.getRowNum());


        hr1.setHeightInPoints(18);
        hr2.setHeightInPoints(18);

        int colA = 0;
        int colMenu = 1;
        int colIng = 2;
        int colD = 3;
        int colE = 4;
        int colF = 5;
        int colG = 6;
        int methodFirstCol = methodFirst;
        int methodLastCol = methodLast;

// English comment: Create base header cells
        for (int c = colMenu; c <= methodLastCol; c++) {
            Cell c1 = hr1.getCell(c);
            if (c1 == null) c1 = hr1.createCell(c);
            c1.setCellStyle(S.header);

            Cell c2 = hr2.getCell(c);
            if (c2 == null) c2 = hr2.createCell(c);
            c2.setCellStyle(S.header);

        }

// A column: blank + vertical merge (2 rows)
        Cell a1 = safeCell(sh, hr1.getRowNum(), colA);
        a1.setCellValue("");
        a1.setCellStyle(S.blankA);

        Cell a2 = safeCell(sh, hr2.getRowNum(), colA);
        a2.setCellValue("");
        a2.setCellStyle(S.blankA);

// Menu (B) vertical merge
        Cell b1 = hr1.getCell(colMenu);
        b1.setCellValue("메뉴명");
        CellRangeAddress mB = new CellRangeAddress(hr1.getRowNum(), hr2.getRowNum(), colMenu, colMenu);
        sh.addMergedRegion(mB);
        setMergedBorder(sh, mB, BorderStyle.THIN);

// Ingredient (C) vertical merge
        Cell c1 = hr1.getCell(colIng);
        c1.setCellValue("식재료명");
        CellRangeAddress mC = new CellRangeAddress(hr1.getRowNum(), hr2.getRowNum(), colIng, colIng);
        sh.addMergedRegion(mC);
        setMergedBorder(sh, mC, BorderStyle.THIN);

// Method (H~L) merge across columns AND 2 rows
        Cell mh = hr1.getCell(methodFirstCol);
        mh.setCellValue("만드는방법");
        CellRangeAddress mMethod = new CellRangeAddress(hr1.getRowNum(), hr2.getRowNum(), methodFirstCol, methodLastCol);
        sh.addMergedRegion(mMethod);
        setMergedBorder(sh, mMethod, BorderStyle.THIN);

// AGE35 vs AGE12 amount/total header layout
        if (mode == OutputMode.AGE35) {
            // Row1: D~E merged => "1인 제공량(g)"
            Cell de = hr1.getCell(colD);
            de.setCellValue("1인 제공량(g)");
            CellRangeAddress mDE = new CellRangeAddress(hr1.getRowNum(), hr1.getRowNum(), colD, colE);
            sh.addMergedRegion(mDE);
            setMergedBorder(sh, mDE, BorderStyle.THIN);

            // Row1: F~G merged => "총 발주량"
            Cell fg = hr1.getCell(colF);
            fg.setCellValue("총 발주량");
            CellRangeAddress mFG = new CellRangeAddress(hr1.getRowNum(), hr1.getRowNum(), colF, colG);
            sh.addMergedRegion(mFG);
            setMergedBorder(sh, mFG, BorderStyle.THIN);

            // Row2 labels + colors
            Cell d2 = hr2.getCell(colD);
            d2.setCellValue("1~2세");
            d2.setCellStyle(S.header);

            Cell e2 = hr2.getCell(colE);
            e2.setCellValue("3~5세");
            e2.setCellStyle(S.header);

            Cell f2 = hr2.getCell(colF);
            f2.setCellValue("1~2세");
            f2.setCellStyle(S.headerSubYellow);

            Cell g2 = hr2.getCell(colG);
            g2.setCellValue("3~5세");
            g2.setCellStyle(S.headerSubPink);
        } else {
            // AGE12: D column => per-serving, E column => total
            Cell d1 = hr1.getCell(colD);
            d1.setCellValue("1인 제공량(g)");
            d1.setCellStyle(S.header);
            Cell e1 = hr1.getCell(colE);
            e1.setCellValue("총 발주량");

            Cell d2 = hr2.getCell(colD);
            d2.setCellValue("1~2세");
            d2.setCellStyle(S.header);
            Cell e2 = hr2.getCell(colE);
            e2.setCellValue("1~2세");
            e2.setCellStyle(S.headerSubYellow);

            // For AGE12, D and E are NOT vertically merged (they are 2-row headers)
            // Other columns already merged above.
        }

// English comment: Ensure borders on D/E/F/G cells (some are inside merged regions)
        for (int cc = colD; cc <= colG; cc++) {
            setBorders(safeCell(sh, hr1.getRowNum(), cc), BorderStyle.THIN);
            setBorders(safeCell(sh, hr2.getRowNum(), cc), BorderStyle.THIN);
        }


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
            SearchResult sr = findBlockInTemplateExactPreferNew(tplNew, tplOld, rawMenu);
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
                        writeDataRowWithFormulas(row, S, menuCell, it.ingredient, it.p12, it.p35, it.t12, it.t35, method, mode);
                    }
                }
            }

            int blockEnd = r - 1;

            // Merge menu column A
            Cell bTop = safeCell(sh, blockStart, 1);
            bTop.setCellStyle(S.bodyCenter);
            if (blockEnd > blockStart) {
                CellRangeAddress bMerge = new CellRangeAddress(blockStart, blockEnd, 1, 1);
                sh.addMergedRegion(bMerge);
                setMergedBorder(sh, bMerge, BorderStyle.THIN);
            } else {
                setBorders(bTop, BorderStyle.THIN);
            }

            for (int rr = blockStart; rr <= blockEnd; rr++) {
                Cell a = safeCell(sh, rr, 0);
                a.setCellValue("");
                a.setCellStyle(S.blankA);
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
                    // English comment: B~G should have borders (menu..totals)
                    for (int bc = 1; bc <= 6; bc++) setBorders(safeCell(sh, br, bc), BorderStyle.THIN);
                }
            } else {
                for (int br = blockStart; br <= blockEnd; br++) {
                    // English comment: B~E should have borders (menu..total)
                    for (int bc = 1; bc <= 4; bc++) setBorders(safeCell(sh, br, bc), BorderStyle.THIN);
                }
            }
        }

        int lastDataRow = r - 1;
        if (lastDataRow >= headerRowIndex) {
            CellRangeAddress outer = new CellRangeAddress(headerRowIndex, lastDataRow, 1, lastColForBlock);
            setMergedBorder(sh, outer, BorderStyle.DOUBLE);

            forceLeftOuterBorderDouble(sh, headerRowIndex, lastDataRow, 1);

            // English comment: Column A must keep width but remain visually blank.
            // English comment: Clear borders and force NO_FILL style to remove template residue.
            int rr = headerRowIndex;
            while (rr <= lastDataRow) {
                Cell a = safeCell(sh, rr, 0);
                a.setCellValue("");
                a.setCellStyle(S.blankA);

                Workbook wb = sh.getWorkbook();
                CellStyle base = S.blankA;

                String key = "COLA_CLEAN|" + System.identityHashCode(base);
                CellStyle derived = getOrCreateStyle(
                        wb,
                        key,
                        base,
                        BorderStyle.NONE, BorderStyle.NONE, BorderStyle.NONE, BorderStyle.NONE,
                        true
                );

                a.setCellStyle(derived);

                rr = rr + 1;
            }
        }

        int gapStart = r;
        for (int gr = 0; gr < GAP_ROWS; gr++) {
            Row gapRow = safeRow(sh, gapStart + gr);
            if (gr == 0) gapRow.setHeightInPoints(80.1f);

            for (int cc = 0; cc <= lastColForBlock; cc++) {
                Cell c = safeCell(sh, gapStart + gr, cc);
                c.setCellValue("");
                if (cc == 0) c.setCellStyle(S.blankA);
                else c.setCellStyle(S.gapBlank); // 이 스타일은 "테두리 NONE + NO_FILL"로 만들어둔 걸 쓰기
            }
        }

        // English comment: Put allergy banner INSIDE the GAP (first gap row).
        addAllergyImageOnRow(sh, gapStart, 1, lastColForBlock);

        r = r + GAP_ROWS;
        return r;
    }

    void writeDataRowWithFormulas(Row row, Styles S, String menu, String ing, String p12, String p35, String t12, String t35, String method, OutputMode mode) {
        int excelRow = row.getRowNum() + 1;
        row.setHeightInPoints(18);

        // English comment: A column should stay blank
        Cell colA = row.createCell(0);
        colA.setCellValue("");
        colA.setCellStyle(S.blankA);

        // B: menu
        Cell menuCell = row.createCell(1);
        menuCell.setCellValue(nz(menu));
        menuCell.setCellStyle(S.bodyCenter);

        // C: ingredient
        Cell ingCell = row.createCell(2);
        ingCell.setCellValue(nz(ing));
        ingCell.setCellStyle(S.bodyCenter);

        if (mode == OutputMode.AGE35) {
            // D: 1~2 (derived from E * 0.65)
            // E: 3~5 (input)
            Cell e = row.createCell(4);

// English comment: In AGE35, D/F/G depend on E being numeric.
// English comment: Parse numeric aggressively; if not numeric, keep text but avoid broken formulas.
            Double p35val = parseNumericOrNull(p35);

            if (p35val != null) {
                e.setCellValue(p35val.doubleValue());
                e.setCellStyle(S.numGeneral);

                Cell d = row.createCell(3);
                d.setCellFormula(String.format("E%d*0.65", excelRow));
                boolean forceOneDecimal = shouldForceOneDecimal(p35);
                if (forceOneDecimal) d.setCellStyle(S.num1dec);
                else d.setCellStyle(S.numGeneral);

                Cell f = row.createCell(5);
                f.setCellFormula(String.format("D%d*$D$2", excelRow));
                f.setCellStyle(S.numGeneral);

                Cell g = row.createCell(6);
                g.setCellFormula(String.format("E%d*$D$3", excelRow));
                g.setCellStyle(S.numGeneral);

            } else {
                // English comment: Keep original text in E, but do NOT set formulas that would produce #VALUE!
                e.setCellValue(nz(p35));
                e.setCellStyle(S.bodyCenter); // English comment: Treat as text.

                Cell d = row.createCell(3);
                d.setCellValue("");
                d.setCellStyle(S.numGeneral);

                Cell f = row.createCell(5);
                f.setCellValue("");
                f.setCellStyle(S.numGeneral);

                Cell g = row.createCell(6);
                g.setCellValue("");
                g.setCellStyle(S.numGeneral);
            }

            // H: method text (actual merge is handled outside)
            Cell h = row.createCell(7);
            h.setCellValue(compactSpaces(nz(method)));
            h.setCellStyle(S.methodMerged);

        } else {
            // AGE12
            // D: 1~2 (input or derived)
            Cell d = row.createCell(3);
            boolean dNum = trySetNumeric(d, p12);
            if (!dNum) {
                Double p35val = parseNumericOrNull(p35);
                if (p35val != null) {
                    double v = p35val * 0.65;
                    d.setCellValue(v);
                    boolean forceOneDecimal = shouldForceOneDecimal(p35);
                    if (forceOneDecimal) d.setCellStyle(S.num1dec);
                    else d.setCellStyle(S.numGeneral);
                } else {
                    d.setCellValue(nz(p12));
                    d.setCellStyle(S.num1dec);
                }
            } else {
                d.setCellStyle(S.num1dec);
            }

            // E: total (no param row -> just equal to D for now)
            Cell e = row.createCell(4);
            e.setCellFormula(String.format("D%d*$D$2", excelRow));
            e.setCellStyle(S.numGeneral);

            // F: method
            Cell f = row.createCell(5);
            f.setCellValue(compactSpaces(nz(method)));
            f.setCellStyle(S.methodMerged);
        }
    }

    boolean shouldForceOneDecimal(String p35) {
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

        SearchResult(boolean f, boolean exact, Block b) {
            this.found = f;
            this.matchedByExact = exact;
            this.block = b;
        }

        static SearchResult miss() {
            return new SearchResult(false, false, null);
        }
    }

    SearchResult findBlockInTemplateExactOnly(Workbook wb, String menuRaw) {
        String keyExact = compactSpacesPreserveAll(applyAlias(menuRaw));
        String keyNorm = normalizeForMatch(keyExact);

        for (int s = wb.getNumberOfSheets() - 1; s >= 0; s--) {
            Sheet sh = wb.getSheetAt(s);
            int lastRow = sh.getLastRowNum();

            int r = 0;
            while (r <= lastRow) {
                String cellMenu = readStringConsideringMerged(sh, r, COL_MENU);
                if (isBlank(cellMenu)) {
                    r++;
                    continue;
                }

                CellRangeAddress menuRange = findMergedRange(sh, r, COL_MENU);
                int first = r;
                int last = r;
                if (menuRange != null) {
                    first = menuRange.getFirstRow();
                    last = menuRange.getLastRow();
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
                String cellNorm = normalizeForMatch(cellExact);

                if (cellExact.equals(keyExact)) return new SearchResult(true, true, readBlock(sh, first, last));
                if (cellNorm.equals(keyNorm)) return new SearchResult(true, false, readBlock(sh, first, last));

                r = last + 1;
            }
        }
        return SearchResult.miss();
    }

    Block readBlock(Sheet sh, int first, int last) {
        String methodTop = "";
        outer:
        for (int rr = first; rr <= last; rr++) {
            for (int cc = COL_METHOD; cc <= COL_METHOD + 4; cc++) {
                String v = readStringConsideringMerged(sh, rr, cc);
                if (!isBlank(v)) {
                    methodTop = v.trim();
                    break outer;
                }
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
        b.items = items;
        return b;
    }

    Path deriveOutXlsxPathFromJson(Path jsonPath, Path outDir) {
        String file = jsonPath.getFileName().toString();
        int dot = file.lastIndexOf('.');
        String stem = (dot > 0) ? file.substring(0, dot) : file;
        String outBase = stem + "_수정";
        return outDir.resolve(outBase + ".xlsx");
    }

    CellRangeAddress findMergedRange(Sheet sh, int r, int c) {
        for (int i = 0; i < sh.getNumMergedRegions(); i++) {
            CellRangeAddress ra = sh.getMergedRegion(i);
            if (ra.isInRange(r, c)) return ra;
        }
        return null;
    }

    String readStringConsideringMerged(Sheet sh, int r, int c) {
        Cell cell = getMergedAnchorCell(sh, r, c);
        if (cell == null) return "";
        return getString(cell).trim();
    }

    Cell getMergedAnchorCell(Sheet sh, int r, int c) {
        for (int i = 0; i < sh.getNumMergedRegions(); i++) {
            CellRangeAddress ra = sh.getMergedRegion(i);
            if (ra.isInRange(r, c)) {
                Row topRow = sh.getRow(ra.getFirstRow());
                if (topRow == null) return null;
                return topRow.getCell(ra.getFirstColumn());
            }
        }
        Row row = sh.getRow(r);
        if (row == null) return null;
        return row.getCell(c);
    }

    String getString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) return cell.getDateCellValue().toString();
                double v = cell.getNumericCellValue();
                long rnd = Math.round(v);
                if (Math.abs(v - rnd) < 1e-9) return String.valueOf(rnd);
                return String.valueOf(v);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    try {
                        return String.valueOf(cell.getNumericCellValue());
                    } catch (Exception ignore) {
                        return "";
                    }
                }
            default:
                return "";
        }
    }

    boolean trySetNumeric(Cell c, String s) {
        if (isBlank(s)) return false;
        try {
            String t = s.replace(",", "").replace("\u00A0", "").replaceAll("[\\u2000-\\u200B\\u202F\\u205F\\u3000]", "").replace('．', '.').replace('。', '.').replace('･', '.').replace('·', '.').trim();
            t = t.replaceAll("[^0-9.\\-]", "");
            t = t.replaceAll("\\.(?=\\.)", "");
            t = t.replaceAll("\\.$", "");
            if (!t.matches("[-]?(\\d+\\.?\\d*|\\d*\\.\\d+)")) return false;

            double v = Double.parseDouble(t);
            c.setCellValue(v);
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    Double parseNumericOrNull(String s) {
        if (isBlank(s)) return null;
        try {
            String t = s.replace(",", "").replace("\u00A0", "").replaceAll("[\\u2000-\\u200B\\u202F\\u205F\\u3000]", "").replace('．', '.').replace('。', '.').replace('･', '.').replace('·', '.').trim();
            t = t.replaceAll("[^0-9.\\-]", "");
            t = t.replaceAll("\\.(?=\\.)", "");
            t = t.replaceAll("\\.$", "");
            if (!t.matches("[-]?(\\d+\\.?\\d*|\\d*\\.\\d+)")) return null;
            return Double.parseDouble(t);
        } catch (Exception e) {
            return null;
        }
    }

    void setBorders(Cell cell, BorderStyle bs) {
        Workbook wb = cell.getSheet().getWorkbook();
        CellStyle base = cell.getCellStyle();

        String key = "BORDERS_ALL|" + System.identityHashCode(base) + "|" + bs.name();
        CellStyle derived = getOrCreateStyle(
                wb,
                key,
                base,
                bs, bs, bs, bs,
                false
        );

        cell.setCellStyle(derived);
    }

    void setMergedBorder(Sheet sh, CellRangeAddress rgn, BorderStyle bs) {
        org.apache.poi.ss.util.RegionUtil.setBorderTop(bs, rgn, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderBottom(bs, rgn, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderLeft(bs, rgn, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderRight(bs, rgn, sh);
    }

    Row safeRow(Sheet sh, int r) {
        Row row = sh.getRow(r);
        if (row == null) row = sh.createRow(r);
        return row;
    }

    Cell safeCell(Sheet sh, int r, int c) {
        Row row = safeRow(sh, r);
        Cell cell = row.getCell(c);
        if (cell == null) cell = row.createCell(c);
        return cell;
    }

    boolean isBlank(String s) {
        return s == null || s.trim().isEmpty();
    }

    String nz(String s) {
        return s == null ? "" : s;
    }

    String compactSpaces(String s) {
        if (s == null) return "";
        return s.replace('\u00A0', ' ').replaceAll("[ \\t]{2,}", " ").trim();
    }

    String compactSpacesPreserveAll(String s) {
        if (s == null) return "";
        return s.replace('\u00A0', ' ').replaceAll("[ \\t]{2,}", " ").trim();
    }

    String applyAlias(String s) {
        if (s == null) return "";
        String t = s;
        for (Map.Entry<String, String> e : ALIAS.entrySet()) {
            t = t.replace(e.getKey(), e.getValue());
        }
        return t;
    }

    String normalizeForMatch(String s) {
        if (s == null) return "";
        String t = applyAlias(s);
        t = t.replaceAll("[\u2460-\u2473①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲]", "");
        t = t.replaceAll("[()\\[\\]{}／/·ㆍ・＆&\\-★☆※•….,;:]", "");
        t = t.replaceAll("\\s+", "");
        return t;
    }

    YearMonth inferYMFromJson(Path jsonPath) {
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
        YearMonth ym2 = tryInferYMFromKoreanName(name);
        if (ym2 != null) return ym2;

// English comment: Last resort: use current year-month.
        return YearMonth.from(LocalDate.now());
    }

    boolean isTailMenu(String s) {
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

    static class Item {
        String ingredient, p12, p35, t12, t35;
    }

    static class Block {
        String displayName;
        String method;
        List<Item> items = new ArrayList<>();
    }

    static class Styles {
        final CellStyle title;
        final CellStyle header;
        final CellStyle headerSubYellow;
        final CellStyle headerSubPink;
        final CellStyle bodyCenter;
        final CellStyle blankA;
        final CellStyle gapBlank;
        final CellStyle methodMerged;
        final CellStyle num1dec;
        final CellStyle numGeneral;
        final CellStyle param;

        Styles(CellStyle title, CellStyle header, CellStyle headerSubYellow, CellStyle headerSubPink, CellStyle bodyCenter, CellStyle blankA, CellStyle gapBlank, CellStyle methodMerged, CellStyle num1dec, CellStyle numGeneral, CellStyle param) {
            this.title = title;
            this.header = header;
            this.headerSubYellow = headerSubYellow;
            this.headerSubPink = headerSubPink;
            this.bodyCenter = bodyCenter;
            this.blankA = blankA;
            this.gapBlank = gapBlank;
            this.methodMerged = methodMerged;
            this.num1dec = num1dec;
            this.numGeneral = numGeneral;
            this.param = param;
        }

        static Styles build(Workbook wb) {
            DataFormat df = wb.createDataFormat();

            Font titleFont = wb.createFont();
            titleFont.setFontName("한컴산뜻돋움");
            titleFont.setFontHeightInPoints((short) 14);
            titleFont.setBold(true);

            Font headerFont = wb.createFont();
            headerFont.setFontName("한컴산뜻돋움");
            headerFont.setFontHeightInPoints((short) 11);
            headerFont.setBold(true);

            Font bodyFont = wb.createFont();
            bodyFont.setFontName("한컴산뜻돋움");
            bodyFont.setFontHeightInPoints((short) 10);

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
            title.setFont(titleFont);

            CellStyle header = wb.createCellStyle();
            header.cloneStyleFrom(base);
            header.setAlignment(HorizontalAlignment.CENTER);
            header.setVerticalAlignment(VerticalAlignment.CENTER);
            header.setWrapText(true);
            header.setFont(headerFont);

            CellStyle headerSubYellow = wb.createCellStyle();
            headerSubYellow.cloneStyleFrom(header);
            // English comment: Approx yellow for #FFFF9F


            CellStyle headerSubPink = wb.createCellStyle();
            headerSubPink.cloneStyleFrom(header);
            // English comment: Approx pink for #FF9B9B

            CellStyle bodyCenter = wb.createCellStyle();
            bodyCenter.cloneStyleFrom(base);
            bodyCenter.setAlignment(HorizontalAlignment.CENTER);
            bodyCenter.setVerticalAlignment(VerticalAlignment.CENTER);
            bodyCenter.setWrapText(true);
            bodyCenter.setFont(bodyFont);

            CellStyle blankA = wb.createCellStyle();
            blankA.cloneStyleFrom(bodyCenter);

// English comment: Column A must have NO borders always.
            blankA.setBorderTop(BorderStyle.NONE);
            blankA.setBorderBottom(BorderStyle.NONE);
            blankA.setBorderLeft(BorderStyle.NONE);
            blankA.setBorderRight(BorderStyle.NONE);

// English comment: Force NO_FILL to wipe template column style residue.
            blankA.setFillPattern(FillPatternType.NO_FILL);
            blankA.setFillForegroundColor(IndexedColors.AUTOMATIC.getIndex());
            blankA.setFillBackgroundColor(IndexedColors.AUTOMATIC.getIndex()); // English comment: extra safety
            blankA.setIndention((short) 0);


            CellStyle gapBlank = wb.createCellStyle();
            gapBlank.cloneStyleFrom(bodyCenter);

// English comment: NO borders for gap rows (prevents stray lines under gapTop row).
            gapBlank.setBorderTop(BorderStyle.NONE);
            gapBlank.setBorderBottom(BorderStyle.NONE);
            gapBlank.setBorderLeft(BorderStyle.NONE);
            gapBlank.setBorderRight(BorderStyle.NONE);

// English comment: Keep it truly blank (no fill).
            gapBlank.setFillPattern(FillPatternType.NO_FILL);
            gapBlank.setFillForegroundColor(IndexedColors.AUTOMATIC.getIndex());
            gapBlank.setFillBackgroundColor(IndexedColors.AUTOMATIC.getIndex());

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

            boolean isXssf = wb instanceof XSSFWorkbook;

            if (isXssf) {
                // #DDEBF7
                ((XSSFCellStyle) title).setFillForegroundColor(new XSSFColor(new java.awt.Color(0xDA, 0xEE, 0xF3), null));
                title.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                // #F2F2F2
                ((XSSFCellStyle) header).setFillForegroundColor(new XSSFColor(new java.awt.Color(0xF2, 0xF2, 0xF2), null));
                header.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                // #FFFF9F
                ((XSSFCellStyle) headerSubYellow).setFillForegroundColor(new XSSFColor(new java.awt.Color(0xFF, 0xFF, 0x9F), null));
                headerSubYellow.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                // #FF9B9B
                ((XSSFCellStyle) headerSubPink).setFillForegroundColor(new XSSFColor(new java.awt.Color(0xFF, 0x9B, 0x9B), null));
                headerSubPink.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            } else {
                title.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
                title.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                header.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                header.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                headerSubYellow.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
                headerSubYellow.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                headerSubPink.setFillForegroundColor(IndexedColors.ROSE.getIndex());
                headerSubPink.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            }
            header.setWrapText(false);

            return new Styles(title, header, headerSubYellow, headerSubPink, bodyCenter, blankA, gapBlank, methodMerged, num1dec, numGeneral, param);
        }
    }

    List<DayPlan> readPlan(Path json) throws IOException {
        ObjectMapper om = new ObjectMapper();
        try (var is = Files.newInputStream(json)) {
            return om.readValue(is, new TypeReference<List<DayPlan>>() {
            });
        }
    }


    // English comment: Read plan list from JSON string (no intermediate JSON file)
    List<DayPlan> readPlanFromString(String jsonString) throws IOException {
        if (jsonString == null) throw new IllegalArgumentException("jsonString is null");
        ObjectMapper om = new ObjectMapper();
        return om.readValue(jsonString, new TypeReference<List<DayPlan>>() {
        });
    }

    private Path doConvertFromJsonString(String jsonString, Path sourceNamePath, Path outXlsxPath) throws Exception {

        ZipSecureFile.setMinInflateRatio(0.0d);
        ZipSecureFile.setMaxFileCount(20000);

        // English comment: Validate inputs
        if (jsonString == null) throw new IllegalArgumentException("jsonString is null");
        if (sourceNamePath == null) throw new IllegalArgumentException("sourceNamePath is null");
        if (outXlsxPath == null) throw new IllegalArgumentException("outXlsxPath is null");


        YearMonth ym = inferYMFromJson(sourceNamePath);
        String titlePrefix = String.format("%d년 %d월", ym.getYear(), ym.getMonthValue());
        String monthPrefix = extractMonthPrefixFromSourceName(sourceNamePath);

        if (outXlsxPath.getParent() != null) Files.createDirectories(outXlsxPath.getParent());

        OutputMode mode = inferModeFromJsonName(sourceNamePath);

        Path baseTplPath = resolveTemplateBase(mode);
        Path tplNewPath = resolveTemplateNew(mode);
        Path tplOldPath = resolveTemplateOld(mode);

        List<DayPlan> plan = readPlanFromString(jsonString);

        try (Workbook out = WorkbookResources.open(baseTplPath);
             Workbook tplNew = WorkbookResources.open(tplNewPath);
             Workbook tplOld = WorkbookResources.open(tplOldPath)) {


            // English comment: Keep only the first sheet as base, then rename to avoid collisions.
            String baseName = out.getSheetName(0);
            for (int i = out.getNumberOfSheets() - 1; i >= 0; i--) {
                String nm = out.getSheetName(i);
                if (!baseName.equals(nm)) out.removeSheetAt(i);
            }
            int baseIdx = 0;


            out.setSheetName(baseIdx, "__BASE_TEMPLATE__");
            Styles S = Styles.build(out);

            // English comment: Use clone-based output sheet
            int DATA_START_ROW;
            if (mode == OutputMode.AGE12) DATA_START_ROW = 5;  // 1-2세: title at Excel row 6
            else DATA_START_ROW = 6;                           // 3-5세: title at Excel row 7

            int keepLastRow = DATA_START_ROW - 1;

            int weekNo = 1;
            Sheet sh = copyTopTemplateArea(out, baseIdx, buildSheetName(monthPrefix, weekNo), keepLastRow);
            sh.setDefaultColumnStyle(0, S.blankA);

            int nextRow = DATA_START_ROW;

            int i = 0;
            while (i < plan.size()) {
                DayPlan d = plan.get(i);
                if (d != null) {
                    nextRow = writeOneDay(sh, S, d, tplNew, tplOld, titlePrefix, mode, nextRow);

                    // English comment: Split sheet after Friday (end of week).
                    if (isFriday(ym, d)) {

                        // English comment: Finish current sheet cleanup.
                        clearColumnAAllRows(sh, S);

                        // English comment: Prepare next week sheet if there are remaining days.
                        if (i + 1 < plan.size()) {
                            weekNo = weekNo + 1;
                            sh = copyTopTemplateArea(out, baseIdx, buildSheetName(monthPrefix, weekNo), keepLastRow);
                            sh.setDefaultColumnStyle(0, S.blankA);
                            nextRow = DATA_START_ROW;
                        }
                    }
                }
                i = i + 1;
            }

// English comment: Ensure last sheet cleanup.
            clearColumnAAllRows(sh, S);

            // English comment: Remove base template sheet
            int baseIdx2 = out.getSheetIndex("__BASE_TEMPLATE__");
            if (baseIdx2 >= 0) out.removeSheetAt(baseIdx2);

            try (OutputStream os = Files.newOutputStream(outXlsxPath)) {
                out.write(os);
            }
        }

        return outXlsxPath;
    }


    Path resolveTemplateBase(OutputMode mode) throws Exception {
        return new TemplateCatalog(appHome).base(true, mode == OutputMode.AGE12);
    }

    // English comment: Resolve lookup "new" template (2026~) if exists
    Path resolveTemplateNew(OutputMode mode) throws Exception {
        return new TemplateCatalog(appHome).newer(true, mode == OutputMode.AGE12);
    }

    // English comment: Resolve lookup "old" template if exists
    Path resolveTemplateOld(OutputMode mode) throws Exception {
        return new TemplateCatalog(appHome).older(true, mode == OutputMode.AGE12);
    }

    SearchResult findBlockInTemplateExactPreferNew(Workbook tplNew, Workbook tplOld, String rawMenu) {
        if (tplNew != null) {
            SearchResult a = findBlockInTemplateExactOnly(tplNew, rawMenu);
            if (a.found) return a;
        }
        if (tplOld != null) {
            SearchResult b = findBlockInTemplateExactOnly(tplOld, rawMenu);
            if (b.found) return b;
        }
        return SearchResult.miss();
    }

    Sheet copyTopTemplateArea(Workbook out, int baseIdx, String newSheetName, int keepLastRow) {
        // 1) Clone sheet (copies column widths, row heights, merged regions, images, drawings)
        Sheet newSh = out.cloneSheet(baseIdx);

        // 2) Rename
        int newIdx = out.getSheetIndex(newSh);
        String finalName = newSheetName;
        int suffix = 1;

        while (true) {
            Sheet existing = out.getSheet(finalName);
            if (existing == null) break;

            int existingIdx = out.getSheetIndex(existing);
            if (existingIdx == newIdx) break;

            finalName = newSheetName + "_" + suffix;
            suffix = suffix + 1;
        }

        out.setSheetName(newIdx, finalName);

        // 3) Remove all rows below keepLastRow (from bottom to top)
        int lastRow = newSh.getLastRowNum();
        int r = lastRow;
        while (r > keepLastRow) {
            Row row = newSh.getRow(r);
            if (row != null) newSh.removeRow(row);
            r = r - 1;
        }

        // 4) Remove merged regions that are fully below keepLastRow,
        //    and also merged regions that cross the boundary (to avoid weird overlaps)
        removeMergedRegionsBelowOrCrossing(newSh, keepLastRow);

        // 5) Remove pictures anchored below keepLastRow (optional but recommended)
        //    If you want to keep all top pictures only.
        removePicturesBelowRow(newSh, keepLastRow);

        return newSh;
    }

    // English comment: Remove merged regions that are below or crossing keepLastRow.
    void removeMergedRegionsBelowOrCrossing(Sheet sh, int keepLastRow) {
        List<Integer> toRemove = new ArrayList<>();
        int i = 0;
        while (i < sh.getNumMergedRegions()) {
            CellRangeAddress ra = sh.getMergedRegion(i);

            boolean fullyBelow = ra.getFirstRow() > keepLastRow;
            boolean crossing = ra.getFirstRow() <= keepLastRow && ra.getLastRow() > keepLastRow;

            if (fullyBelow || crossing) toRemove.add(i);
            i = i + 1;
        }

        Collections.reverse(toRemove);
        for (int idx : toRemove) sh.removeMergedRegion(idx);
    }

    // English comment: Remove pictures that start below keepLastRow (XSSF only).
    void removePicturesBelowRow(Sheet sh, int keepLastRow) {
        if (!(sh instanceof org.apache.poi.xssf.usermodel.XSSFSheet)) return;

        org.apache.poi.xssf.usermodel.XSSFSheet xs = (org.apache.poi.xssf.usermodel.XSSFSheet) sh;
        org.apache.poi.xssf.usermodel.XSSFDrawing drawing = xs.getDrawingPatriarch();
        if (drawing == null) return;

        org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing ct = drawing.getCTDrawing();
        if (ct == null) return;

        // English comment: Remove anchors whose top row is below keepLastRow.
        // English comment: Iterate from end to start to avoid index shift while removing.
        int i;

        i = ct.sizeOfTwoCellAnchorArray() - 1;
        while (i >= 0) {
            org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor a = ct.getTwoCellAnchorArray(i);
            int row1 = -1;
            if (a != null && a.getFrom() != null) row1 = a.getFrom().getRow();
            if (row1 > keepLastRow) ct.removeTwoCellAnchor(i);
            i = i - 1;
        }

        i = ct.sizeOfOneCellAnchorArray() - 1;
        while (i >= 0) {
            org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTOneCellAnchor a = ct.getOneCellAnchorArray(i);
            int row1 = -1;
            if (a != null && a.getFrom() != null) row1 = a.getFrom().getRow();
            if (row1 > keepLastRow) ct.removeOneCellAnchor(i);
            i = i - 1;
        }

        i = ct.sizeOfAbsoluteAnchorArray() - 1;
        while (i >= 0) {
            // English comment: Absolute anchors are rare in templates; remove all to be safe if present.
            ct.removeAbsoluteAnchor(i);
            i = i - 1;
        }
    }

    void clearBordersInColumnA(Sheet sh, int firstRow, int lastRow) {
        int rr = firstRow;
        while (rr <= lastRow) {
            Cell a = safeCell(sh, rr, 0);

            CellStyle cs = sh.getWorkbook().createCellStyle();
            cs.cloneStyleFrom(a.getCellStyle());

            cs.setBorderTop(BorderStyle.NONE);
            cs.setBorderBottom(BorderStyle.NONE);
            cs.setBorderLeft(BorderStyle.NONE);
            cs.setBorderRight(BorderStyle.NONE);

            // English comment: Keep it truly blank (no fill).
            cs.setFillPattern(FillPatternType.NO_FILL);
            cs.setFillForegroundColor(IndexedColors.AUTOMATIC.getIndex());
            cs.setFillBackgroundColor(IndexedColors.AUTOMATIC.getIndex());

            a.setCellStyle(cs);
            rr = rr + 1;
        }
    }

    void clearColumnAAllRows(Sheet sh, Styles S) {
        int last = sh.getLastRowNum();
        int r = 0;
        while (r <= last) {
            Row row = sh.getRow(r);
            if (row == null) {
                r = r + 1;
                continue;
            }

            Cell a = row.getCell(0);
            if (a == null) a = row.createCell(0);

            a.setCellValue("");
            a.setCellStyle(S.blankA);

            // English comment: Remove any borders that leaked from merged/outer borders.
            CellStyle cs = sh.getWorkbook().createCellStyle();
            cs.cloneStyleFrom(a.getCellStyle());
            cs.setBorderTop(BorderStyle.NONE);
            cs.setBorderBottom(BorderStyle.NONE);
            cs.setBorderLeft(BorderStyle.NONE);
            cs.setBorderRight(BorderStyle.NONE);

            // English comment: Extra safety - keep A column truly blank even if template had fills.
            cs.setFillPattern(FillPatternType.NO_FILL);
            cs.setFillForegroundColor(IndexedColors.AUTOMATIC.getIndex());
            cs.setFillBackgroundColor(IndexedColors.AUTOMATIC.getIndex());

            a.setCellStyle(cs);

            r = r + 1;
        }
    }

    Path resolveAllergyImagePath() {
        Path appHome = getAppHomeDir();
        Path p1 = appHome.resolve("input").resolve("allergy.png");
        if (Files.exists(p1)) return p1;

        Path p2 = Paths.get("").toAbsolutePath().resolve("input").resolve("allergy.png");
        if (Files.exists(p2)) return p2;

        return null;
    }

    void addAllergyImageOnRow(Sheet sh, int rowIndex, int firstCol, int lastCol) {
        try {
            Path imgPath = resolveAllergyImagePath();
            if (imgPath == null) return;

            byte[] imgBytes = Files.readAllBytes(imgPath);
            int picIdx = sh.getWorkbook().addPicture(imgBytes, Workbook.PICTURE_TYPE_PNG);

            Row imgRow = sh.getRow(rowIndex);
            if (imgRow == null) imgRow = sh.createRow(rowIndex);

            Drawing<?> drawing = sh.createDrawingPatriarch();
            CreationHelper helper = sh.getWorkbook().getCreationHelper();

            ClientAnchor anchor = helper.createClientAnchor();
            anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);

            // English comment: Span the same width as the day table (B..LAST_COL).
            anchor.setCol1(firstCol);
            anchor.setCol2(lastCol + 1);   // exclusive
            anchor.setRow1(rowIndex);
            anchor.setRow2(rowIndex + 1);  // exclusive

            anchor.setDx1(0);
            anchor.setDy1(0);
            anchor.setDx2(0);
            anchor.setDy2(0);

            org.apache.poi.xssf.usermodel.XSSFPicture pic =
                    (org.apache.poi.xssf.usermodel.XSSFPicture) drawing.createPicture(anchor, picIdx);

            // English comment: Fit image into the anchor cell range.
            pic.resize(1.00);

            clearBordersInColumnA(sh, rowIndex, rowIndex);

        } catch (Exception ignore) {
            // English comment: Ignore image errors to avoid breaking XLSX generation.
        }
    }


    void forceLeftOuterBorderDouble(Sheet sh, int firstRow, int lastRow, int leftCol) {
        int rr = firstRow;
        while (rr <= lastRow) {
            Cell c = safeCell(sh, rr, leftCol);

            CellStyle cs = sh.getWorkbook().createCellStyle();
            cs.cloneStyleFrom(c.getCellStyle());

            cs.setBorderLeft(BorderStyle.DOUBLE);

            c.setCellStyle(cs);
            rr = rr + 1;
        }
    }

    boolean isFriday(String weekday) {
        if (weekday == null) return false;

        String t = weekday.trim();
        if (t.length() == 0) return false;

        // English comment: Accept common Korean labels and also cases like "금(요일)".
        if (t.equals("금") || t.startsWith("금")) return true;
        if (t.toLowerCase(Locale.ROOT).contains("fri")) return true;

        return false;
    }

    boolean isFriday(YearMonth ym, DayPlan d) {
        if (ym == null) return false;
        if (d == null) return false;

        // English comment: Always prefer computed day-of-week from YearMonth + day-of-month
        // English comment: to avoid JSON weekday text issues (e.g., wrong labels in time-extension plan).
        try {
            LocalDate dt = ym.atDay(d.date);
            return dt.getDayOfWeek() == DayOfWeek.FRIDAY;
        } catch (Exception ignore) {
            // English comment: Fallback only when date is invalid.
        }

        // English comment: Fallback to text-based check (last resort).
        return isFriday(d.weekday);
    }

    void removeHorizontalMergesInColumnA(Sheet sh, int row1, int row2) {
        if (sh == null) return;

        List<Integer> toRemove = new ArrayList<>();
        int i = 0;
        while (i < sh.getNumMergedRegions()) {
            CellRangeAddress ra = sh.getMergedRegion(i);

            boolean rowOverlap = ra.getFirstRow() <= row2 && ra.getLastRow() >= row1;
            boolean includesA = ra.getFirstColumn() == 0;
            boolean horizontal = ra.getLastColumn() > 0;

            if (rowOverlap && includesA && horizontal) toRemove.add(i);

            i = i + 1;
        }

        Collections.reverse(toRemove);
        for (int idx : toRemove) sh.removeMergedRegion(idx);
    }

    private String extractMonthPrefixFromSourceName(Path sourceNamePath) {
        // English comment: Parse "2026년 1월 ..." from source file name.
        String name = sourceNamePath.getFileName().toString();

        // English comment: Remove extension for cleaner matching.
        int dot = name.lastIndexOf('.');
        if (dot > 0) {
            name = name.substring(0, dot);
        }

        Pattern p = Pattern.compile("(\\d{4})\\s*년\\s*(\\d{1,2})\\s*월");
        Matcher m = p.matcher(name);

        if (!m.find()) {
            // English comment: Fallback prefix when filename doesn't match.
            return "00.0월";
        }

        int year = Integer.parseInt(m.group(1));
        int month = Integer.parseInt(m.group(2));

        // English comment: Convert 2026 -> 26
        int yy = year % 100;

        // English comment: Build "26.1월"
        return yy + "." + month + "월";
    }

    private String buildSheetName(String monthPrefix, int weekIndex) {
        String wk;
        if (weekIndex == 1) wk = "첫째";
        else if (weekIndex == 2) wk = "둘째";
        else if (weekIndex == 3) wk = "셋째";
        else if (weekIndex == 4) wk = "넷째";
        else if (weekIndex == 5) wk = "다섯째";
        else if (weekIndex == 6) wk = "여섯째";
        else wk = String.valueOf(weekIndex) + "째";

        String sheetName = monthPrefix.replace(".", ". ") + " " + wk + "주";

        // English comment: Excel sheet name length limit is 31.
        if (sheetName.length() > 31) {
            sheetName = sheetName.substring(0, 31);
        }

        // English comment: Replace forbidden characters for Excel sheet names.
        sheetName = sheetName.replace("/", "_");
        sheetName = sheetName.replace("\\", "_");
        sheetName = sheetName.replace("[", "(");
        sheetName = sheetName.replace("]", ")");
        sheetName = sheetName.replace(":", "-");
        sheetName = sheetName.replace("*", "_");

        return sheetName;
    }


    CellStyle getOrCreateStyle(Workbook wb, String key, CellStyle base, BorderStyle top, BorderStyle bottom, BorderStyle left, BorderStyle right,
                                      boolean noFill) {

        CellStyle cached = STYLE_CACHE.get(key);
        if (cached != null) return cached;

        CellStyle cs = wb.createCellStyle();
        cs.cloneStyleFrom(base);

        if (top != null) cs.setBorderTop(top);
        if (bottom != null) cs.setBorderBottom(bottom);
        if (left != null) cs.setBorderLeft(left);
        if (right != null) cs.setBorderRight(right);

        if (noFill) {
            cs.setFillPattern(FillPatternType.NO_FILL);
            cs.setFillForegroundColor(IndexedColors.AUTOMATIC.getIndex());
            cs.setFillBackgroundColor(IndexedColors.AUTOMATIC.getIndex());
        }

        STYLE_CACHE.put(key, cs);
        return cs;
    }

    YearMonth tryInferYMFromKoreanName(String filename) {
        if (filename == null) return null;

        String name = filename;

        // English comment: Remove extension for cleaner matching.
        int dot = name.lastIndexOf('.');
        if (dot > 0) name = name.substring(0, dot);

        // English comment: Match "2026년 1월" (allow spaces)
        Matcher m = Pattern.compile("(20\\d{2})\\s*년\\s*(\\d{1,2})\\s*월").matcher(name);
        if (!m.find()) return null;

        int y = Integer.parseInt(m.group(1));
        int mo = Integer.parseInt(m.group(2));

        if (mo < 1) mo = 1;
        if (mo > 12) mo = 12;

        return YearMonth.of(y, mo);
    }

    // New: JSON -> Excel with explicit paths
    private Path doConvert(Path jsonPlan, Path outXlsxPath) throws Exception {
        return doConvertFromJsonString(Files.readString(jsonPlan, java.nio.charset.StandardCharsets.UTF_8), jsonPlan, outXlsxPath);
    }

    public static Path convertFromJsonString(String json, Path source, Path output, Path appHome) throws Exception {
        return new JsonToExcel(appHome).doConvertFromJsonString(json, source, output);
    }

    public static void main(String[] args) throws Exception {
        new JsonToExcel().runMain(args);
    }

    public static Path convertFromJsonString(String jsonString, Path sourceNamePath, Path outXlsxPath) throws Exception {
        return new JsonToExcel().doConvertFromJsonString(jsonString, sourceNamePath, outXlsxPath);
    }

    public static Path convert(Path jsonPlan, Path outXlsxPath) throws Exception {
        return new JsonToExcel().doConvert(jsonPlan, outXlsxPath);
    }

}
