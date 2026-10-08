// (full file) src/main/java/org/example/JsonToExcelGeneral.java
package org.example;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.Color;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.*;
import java.time.LocalDate;
import java.time.YearMonth;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * JSON (일반형) → 조리지시서 XLSX
 * - Supports two template kinds: P12(만1-2세) and P35(만3-5세)
 * - First morningCount menus[] = 오전간식, remaining menus[] = 점심, pmDesert[] = 오후간식
 * - Dedup consecutive identical menus per section
 * - Kimchi-like menus are moved to bottom of lunch block (just above PM block)
 * - Search: exact match only (allergen-insensitive compare). If not found, keep ONLY menu name (with allergens) and formulas; others blank.
 *
 * New:
 * - Split sheets by week (General ends on Saturday)
 * - Birthday plans go to a dedicated sheet
 * - Sheet naming based on inferred YearMonth from input file name
 */
public class JsonToExcelGeneral {
    private final Path appHome;

    private JsonToExcelGeneral() { this(TemplateCatalog.defaultHome()); }

    private JsonToExcelGeneral(Path appHome) { this.appHome = appHome.toAbsolutePath(); }


    /* ===== Paths ===== */
    // English comment: Default is only fallback when no args are provided
    static final Path JSON_PLAN = Paths.get("output/26.01. 만3-5세 일반형.json");

    // English comment: Template filenames (resolved under app.home/input first)


    static final Set<String> NO_SCALE_MENUS_NORM = new HashSet<>(Arrays.asList(
            normalizeForMatch("우유"),
            normalizeForMatch("두유"),
            normalizeForMatch("아기용치즈"),
            normalizeForMatch("호상요구르트"),
            normalizeForMatch("액상요구르트")
    ));

    /* ===== Template kinds ===== */
    enum TplKind { P12, P35 }

    final Map<String, CellStyle> BORDER_STYLE_CACHE = new HashMap<>();

    /* ===== Week Sheet Name ===== */
    String weekNameKorean(int weekIndex) {
        if (weekIndex == 1) return "첫째주";
        if (weekIndex == 2) return "둘째주";
        if (weekIndex == 3) return "셋째주";
        if (weekIndex == 4) return "넷째주";
        if (weekIndex == 5) return "다섯째주";
        return weekIndex + "주차";
    }

    static final int OUT_COL_OFFSET = 1;

    String makeWeekSheetName(YearMonth ym, int weekIndex) {
        // English comment: Month should not be zero-padded. Format: "26. 1월 첫째주"
        return String.format(
                "%02d. %d월 %s",
                ym.getYear() % 100,
                ym.getMonthValue(),
                weekNameKorean(weekIndex)
        );
    }

    String makeBirthdaySheetName(YearMonth ym) {
        return String.format(
                "%02d. %d월 ★birthday",
                ym.getYear() % 100,
                ym.getMonthValue()
        );
    }

    boolean isWeekEndGeneral(String weekday) {
        if (weekday == null) return false;
        String w = weekday.trim();
        if (w.length() == 0) return false;
        return w.contains("토");
    }

    /* ===== Template column mapping (0-based indexes) ===== */
    static class TemplateCols {
        final int colMenu, colIng, colP12, colP35, colT12, colT35, colMethodStart, colMethodEndIncl;
        TemplateCols(int menu, int ing, int p12, int p35, int t12, int t35, int mStart, int mEnd) {
            colMenu = menu; colIng = ing; colP12 = p12; colP35 = p35; colT12 = t12; colT35 = t35;
            colMethodStart = mStart; colMethodEndIncl = mEnd;
        }
    }

    TplKind detectTplKind(Path tplPath) {
        String name = tplPath.getFileName().toString().toLowerCase(Locale.ROOT);
        if (name.contains("만1-2") || name.contains("1-2세")) return TplKind.P12;
        return TplKind.P35;
    }

    TemplateCols colsOf(TplKind kind) {
        if (kind == TplKind.P12) {
            // 템플릿에서 A열이 공백 → 모든 참조 열을 +1 이동
            // C=menu, D=ing, E=p12, (no p35), F=t12, (no t35), I..M method
            return new TemplateCols(
                    2,
                    3,
                    4,
                    -1,
                    5,
                    -1,
                    8,
                    12
            );
        } else {
            // P35: C=menu, D=ing, E=p12, F=p35, G=t12, H=t35, I..M method
            return new TemplateCols(
                    2,
                    3,
                    4,
                    5,
                    6,
                    7,
                    8,
                    12
            );
        }
    }

    /* ===== Template block model ===== */
    static class Item { String ingredient, p12, p35, t12, t35; }
    static class Block {
        String displayName;
        String method;
        List<Item> items = new ArrayList<>();
    }

    /* ===== JSON model ===== */
    @JsonIgnoreProperties(ignoreUnknown = true)
    public static class DayPlan {
        public Integer date;
        public String dateStr;
        public String weekday;
        public List<String> menus;
        public List<String> pmDesert;
        public Integer morningCount;
    }

    /* ===== Styles ===== */
    static class Styles {
        final CellStyle title, header, bodyCenter, methodMerged, num1;
        final CellStyle banner;
        final CellStyle headerNoFill;
        final CellStyle labelAm, labelLunch, labelPm;
        final CellStyle dec1;

        // English comment: Kids-count input cells (E2/E3) with specific fill colors
        final CellStyle kidsCnt12Fill; // #FFFF9F
        final CellStyle kidsCnt35Fill; // #FF9B9B
        final CellStyle headerKidsFill12; // #FFFF9F
        final CellStyle headerKidsFill35;

        Styles(CellStyle title, CellStyle banner,CellStyle header,CellStyle headerNoFill, CellStyle bodyCenter,
                CellStyle methodMerged, CellStyle num1, CellStyle dec1,
                CellStyle labelAm, CellStyle labelLunch, CellStyle labelPm,
                CellStyle kidsCnt12Fill, CellStyle kidsCnt35Fill, CellStyle headerKidsFill12,
                CellStyle headerKidsFill35) {
            this.title = title; this.header = header; this.bodyCenter = bodyCenter;
            this.methodMerged = methodMerged; this.num1 = num1; this.dec1 = dec1;
            this.labelAm = labelAm; this.labelLunch = labelLunch; this.labelPm = labelPm;
            this.kidsCnt12Fill = kidsCnt12Fill;
            this.kidsCnt35Fill = kidsCnt35Fill;
            this.headerKidsFill12 = headerKidsFill12;
            this.headerKidsFill35 = headerKidsFill35;
            this.headerNoFill = headerNoFill;
            this.banner = banner;
        }

        static Styles build(Workbook wb){
            DataFormat df = wb.createDataFormat();

            Font titleFont = wb.createFont();
            titleFont.setFontName("한컴산뜻돋움");
            titleFont.setFontHeightInPoints((short) 14);
            titleFont.setBold(true);

            Font bodyFont = wb.createFont();
            bodyFont.setFontName("한컴산뜻돋움");
            bodyFont.setFontHeightInPoints((short)10);

            Font headerFont = wb.createFont();
            headerFont.setFontName("한컴산뜻돋움");
            headerFont.setBold(true);
            headerFont.setFontHeightInPoints((short)11);

            Font bannerFont = wb.createFont();
            bannerFont.setFontName("한컴산뜻돋움");
            bannerFont.setBold(true);
            bannerFont.setFontHeightInPoints((short)40);

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

            title.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            if (wb instanceof XSSFWorkbook) {
                XSSFCellStyle xsTitle = (XSSFCellStyle) title;
                xsTitle.setFillForegroundColor(new XSSFColor(new Color(0xDA, 0xEE, 0xF3), null));
            } else {
                title.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
            }

            CellStyle banner = wb.createCellStyle();
            banner.cloneStyleFrom(base);
            banner.setAlignment(HorizontalAlignment.CENTER);
            banner.setVerticalAlignment(VerticalAlignment.CENTER);
            banner.setWrapText(false);
            banner.setFont(bannerFont);

            CellStyle header = wb.createCellStyle();
            header.cloneStyleFrom(base);
            header.setAlignment(HorizontalAlignment.CENTER);
            header.setVerticalAlignment(VerticalAlignment.CENTER);
            header.setWrapText(true);
            header.setFont(headerFont);

            CellStyle headerNoFill = wb.createCellStyle();
            headerNoFill.cloneStyleFrom(base);
            headerNoFill.setAlignment(HorizontalAlignment.CENTER);
            headerNoFill.setVerticalAlignment(VerticalAlignment.CENTER);
            headerNoFill.setWrapText(true);
            headerNoFill.setFont(headerFont);

            header.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            if (wb instanceof XSSFWorkbook) {
                XSSFCellStyle xsHeader = (XSSFCellStyle) header;
                xsHeader.setFillForegroundColor(new XSSFColor(new Color(0xF2, 0xF2, 0xF2), null));
            } else {
                header.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            }

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

            CellStyle num1 = wb.createCellStyle();
            num1.cloneStyleFrom(bodyCenter);
            num1.setDataFormat(df.getFormat("General"));

            CellStyle dec1 = wb.createCellStyle();
            dec1.cloneStyleFrom(bodyCenter);
            dec1.setDataFormat(df.getFormat("0.0"));

            CellStyle labelBase = wb.createCellStyle();
            labelBase.cloneStyleFrom(bodyCenter);
            labelBase.setFont(headerFont);

            CellStyle labelAm    = makeFill(wb, labelBase, 0xE7, 0xD8, 0xEC);
            CellStyle labelLunch = makeFill(wb, labelBase, 0xFF, 0xDB, 0xAB);
            CellStyle labelPm    = makeFill(wb, labelBase, 0xD8, 0xE4, 0xBC);

            CellStyle kidsCnt12Fill = makeFill(wb, num1, 0xFF, 0xFF, 0x9F); // #FFFF9F
            CellStyle kidsCnt35Fill = makeFill(wb, num1, 0xFF, 0x9B, 0x9B); // #FF9B9B
            CellStyle headerKidsFill12 = makeFill(wb, header, 0xFF, 0xFF, 0x9F); // #FFFF9F
            CellStyle headerKidsFill35 = makeFill(wb, header, 0xFF, 0x9B, 0x9B); // #FF9B9B

            header.setWrapText(false);
            labelAm.setWrapText(false);
            labelPm.setWrapText(false);

            return new Styles(title, banner, header, headerNoFill, bodyCenter, methodMerged, num1, dec1,
                    labelAm, labelLunch, labelPm,
                    kidsCnt12Fill, kidsCnt35Fill, headerKidsFill12, headerKidsFill35);
        }

        // English comment: Solid foreground fill with exact RGB (XSSF)
        static CellStyle makeFill(Workbook wb, CellStyle base, int r, int g, int b){
            CellStyle s = wb.createCellStyle();
            s.cloneStyleFrom(base);
            s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            if (wb instanceof XSSFWorkbook) {
                XSSFCellStyle xs = (XSSFCellStyle) s;
                xs.setFillForegroundColor(new XSSFColor(new Color(r, g, b), null));
            } else {
                s.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            }
            return s;
        }
    }

    /* ===== Main ===== */
    private void runMain(String[] args) throws Exception {
        ZipSecureFile.setMinInflateRatio(0.0d);
        ZipSecureFile.setMaxFileCount(20000);

        Path jsonPlan;
        Path outXlsx;

        if (args != null && args.length >= 1 && args[0] != null && args[0].trim().length() > 0) {
            jsonPlan = Paths.get(args[0]).toAbsolutePath();
        } else {
            jsonPlan = JSON_PLAN.toAbsolutePath();
        }

        if (args != null && args.length >= 2 && args[1] != null && args[1].trim().length() > 0) {
            outXlsx = Paths.get(args[1]).toAbsolutePath();
        } else {
            outXlsx = deriveOutXlsxPathFromJson(jsonPlan, Paths.get("output")).toAbsolutePath();
        }

        doConvert(jsonPlan, outXlsx);
        System.out.println("DONE → " + outXlsx.toAbsolutePath());
    }

    // English comment: Resolve template path for general mode (installed app location first)
    Path resolveGeneralTemplate(Path jsonPlan) throws Exception {
        return new TemplateCatalog(appHome).base(false, SourceMetadata.from(jsonPlan).generalAge12());
    }

    Path resolveGeneralTemplateNew(Path jsonPlan) throws Exception {
        return new TemplateCatalog(appHome).newer(false, SourceMetadata.from(jsonPlan).generalAge12());
    }

    Path resolveGeneralTemplateOld(Path jsonPlan) throws Exception {
        return new TemplateCatalog(appHome).older(false, SourceMetadata.from(jsonPlan).generalAge12());
    }

    // English comment: Get app home directory for packaged app
    Path getAppHomeDir() { return appHome; }

    /* ===== One day block (stacked by rows) ===== */
    int writeOneDay(Sheet sh, Styles S, DayPlan d,
            Workbook tplNew, Workbook tplOld, TemplateCols T, TplKind kind,
            String titlePrefix, int startRow, boolean isBirthday) {

        int METHOD_FIRST;
        if (kind == TplKind.P35) METHOD_FIRST = 7 + OUT_COL_OFFSET;
        else METHOD_FIRST = 5 + OUT_COL_OFFSET;

        final int LAST_COL = METHOD_FIRST + 4;

        int[] widths35 = { 3500, 5500, 5000, 3500, 3500, 3500, 3500,4500, 4500, 4500, 4500, 4500 };
        int[] widths12 = { 3500, 5500, 5000, 3500, 3500, 6500, 1,1,1,1,1,1 };

        int r = startRow;
        if (startRow > 0) sh.setRowBreak(startRow - 1);

        List<String> am = new ArrayList<>();
        List<String> lunchOrig = new ArrayList<>();

        if (d.menus != null && !d.menus.isEmpty()) {
            int totalMenus = d.menus.size();
            int morningCnt;

            if (d.morningCount != null && d.morningCount > 0) morningCnt = d.morningCount;
            else morningCnt = 1;

            if (morningCnt > totalMenus) morningCnt = totalMenus;

            int i = 0;
            while (i < morningCnt) { am.add(d.menus.get(i)); i = i + 1; }
            while (i < totalMenus) { lunchOrig.add(d.menus.get(i)); i = i + 1; }
        }

        List<String> pm;
        if (d.pmDesert == null) pm = new ArrayList<>();
        else pm = d.pmDesert;

        // Move kimchi to bottom of lunch
        List<String> lunchMain = new ArrayList<>();
        List<String> lunchKimchi = new ArrayList<>();
        for (String m : lunchOrig) {
            if (isKimchi(m)) lunchKimchi.add(m);
            else lunchMain.add(m);
        }
        List<String> lunch = new ArrayList<>();
        lunch.addAll(lunchMain);
        lunch.addAll(lunchKimchi);

        // Dedup consecutive identical menus
        am    = compressConsecutive(am);
        lunch = compressConsecutive(lunch);
        pm    = compressConsecutive(pm);

        if (am.isEmpty() && lunch.isEmpty() && pm.isEmpty()) {
            return startRow;
        }

        // Title
        Row tr = sh.createRow(r++);
        tr.setHeightInPoints(21);
        Cell t0 = tr.createCell(OUT_COL_OFFSET);

        String titleText;
        if (isBirthday) titleText = "★Birthday식단";
        else titleText = String.format("%s %d일 (%s)", titlePrefix, d.date, nz(d.weekday));

        t0.setCellValue(titleText);
        t0.setCellStyle(S.title);
        addMergeSafe(sh, new CellRangeAddress(tr.getRowNum(), tr.getRowNum(), OUT_COL_OFFSET, LAST_COL));

        CellRangeAddress titleOuter = new CellRangeAddress(tr.getRowNum(), tr.getRowNum(), OUT_COL_OFFSET, LAST_COL);
        setMergedBorder(sh, titleOuter, BorderStyle.DOUBLE);

        org.apache.poi.ss.util.RegionUtil.setBorderBottom(BorderStyle.THIN, titleOuter, sh);

        // Header
        int headerRowIndex;

        if (kind == TplKind.P12) {

            // English comment: 2-row header for P12 with vertical merges on B/C/D and method block.
            Row hrTop = sh.createRow(r++);
            Row hrBot = sh.createRow(r++);

            hrTop.setHeightInPoints(22);
            hrBot.setHeightInPoints(22);

            // Column mapping with OUT_COL_OFFSET (A blank)
            int colB = 0 + OUT_COL_OFFSET; // 구분
            int colC = 1 + OUT_COL_OFFSET; // 메뉴명
            int colD = 2 + OUT_COL_OFFSET; // 식재료명
            int colE = 3 + OUT_COL_OFFSET; // 1인 제공량
            int colF = 4 + OUT_COL_OFFSET; // 총 발주량(요청상 top blank)
            int methodFirst = METHOD_FIRST;
            int methodLast  = LAST_COL;
            // --- Top row labels ---
            Cell b1 = hrTop.createCell(colB); b1.setCellValue("구분");     b1.setCellStyle(S.header);
            Cell c1 = hrTop.createCell(colC); c1.setCellValue("메뉴명");   c1.setCellStyle(S.header);
            Cell d1 = hrTop.createCell(colD); d1.setCellValue("식재료명"); d1.setCellStyle(S.header);

            Cell e1 = hrTop.createCell(colE); e1.setCellValue("1인 제공량(g)"); e1.setCellStyle(S.header);

            Cell f1 = hrTop.createCell(colF); f1.setCellValue("총 발주량"); f1.setCellStyle(S.header);

            // Method top label (merged across method block)
            Cell m1 = hrTop.createCell(methodFirst); m1.setCellValue("만드는방법"); m1.setCellStyle(S.header);

            // --- Bottom row labels ---
            Cell b2 = hrBot.createCell(colB); b2.setCellStyle(S.header);
            Cell c2 = hrBot.createCell(colC); c2.setCellStyle(S.headerNoFill);
            Cell d2 = hrBot.createCell(colD); d2.setCellStyle(S.headerNoFill);

            Cell e2 = hrBot.createCell(colE); e2.setCellValue("1~2세"); e2.setCellStyle(S.header);
            Cell f2 = hrBot.createCell(colF); f2.setCellValue("1~2세"); f2.setCellStyle(S.headerKidsFill12); // #FFFF9F

            Cell m2 = hrBot.createCell(methodFirst); m2.setCellStyle(S.header);

            // --- Merges: B/C/D vertical, method vertical+horizontal ---
            addMergeSafe(sh, new CellRangeAddress(hrTop.getRowNum(), hrBot.getRowNum(), colB, colB));
            addMergeSafe(sh, new CellRangeAddress(hrTop.getRowNum(), hrBot.getRowNum(), colC, colC));
            addMergeSafe(sh, new CellRangeAddress(hrTop.getRowNum(), hrBot.getRowNum(), colD, colD));

            // Method: merge across columns and rows
            addMergeSafe(sh, new CellRangeAddress(hrTop.getRowNum(), hrTop.getRowNum(), methodFirst, methodLast));
            addMergeSafe(sh, new CellRangeAddress(hrBot.getRowNum(), hrBot.getRowNum(), methodFirst, methodLast));
            addMergeSafe(sh, new CellRangeAddress(hrTop.getRowNum(), hrBot.getRowNum(), methodFirst, methodLast));

        }  else {

            // English comment: 2-row header for P35 with grouped columns (serving + totals) and vertical merges on B/C/D and method block.
            Row hrTop = sh.createRow(r++);
            Row hrBot = sh.createRow(r++);
            headerRowIndex = hrTop.getRowNum();

            hrTop.setHeightInPoints(18);
            hrBot.setHeightInPoints(18);

            // Column mapping with OUT_COL_OFFSET (A blank)
            int colB = 0 + OUT_COL_OFFSET; // 구분
            int colC = 1 + OUT_COL_OFFSET; // 메뉴명
            int colD = 2 + OUT_COL_OFFSET; // 식재료명

            int colE = 3 + OUT_COL_OFFSET; // 1인 제공량(1~2)
            int colF = 4 + OUT_COL_OFFSET; // 1인 제공량(3~5)

            int colG = 5 + OUT_COL_OFFSET; // 총 발주량(1~2)
            int colH = 6 + OUT_COL_OFFSET; // 총 발주량(3~5)

            int methodFirst = METHOD_FIRST;
            int methodLast  = LAST_COL;

            // --- Top row labels ---
            Cell b1 = hrTop.createCell(colB); b1.setCellValue("구분");     b1.setCellStyle(S.header);
            Cell c1 = hrTop.createCell(colC); c1.setCellValue("메뉴명");   c1.setCellStyle(S.header);
            Cell d1 = hrTop.createCell(colD); d1.setCellValue("식재료명"); d1.setCellStyle(S.header);

            // Group headers
            Cell e1 = hrTop.createCell(colE); e1.setCellValue("1인 제공량(g)"); e1.setCellStyle(S.header);
            Cell g1 = hrTop.createCell(colG); g1.setCellValue("총 발주량");     g1.setCellStyle(S.header);

            // Method top label
            Cell m1 = hrTop.createCell(methodFirst); m1.setCellValue("만드는방법"); m1.setCellStyle(S.header);

            // --- Bottom row labels ---
            Cell b2 = hrBot.createCell(colB); b2.setCellStyle(S.header);
            Cell c2 = hrBot.createCell(colC); c2.setCellStyle(S.header);
            Cell d2 = hrBot.createCell(colD); d2.setCellStyle(S.header);

            Cell e2 = hrBot.createCell(colE); e2.setCellValue("1~2세"); e2.setCellStyle(S.header);
            Cell f2 = hrBot.createCell(colF); f2.setCellValue("3~5세"); f2.setCellStyle(S.header);

            Cell g2 = hrBot.createCell(colG); g2.setCellValue("1~2세"); g2.setCellStyle(S.headerKidsFill12);   // #FFFF9F
            Cell h2 = hrBot.createCell(colH); h2.setCellValue("3~5세"); h2.setCellStyle(S.headerKidsFill35); // #FF9B9B

            Cell m2 = hrBot.createCell(methodFirst); m2.setCellStyle(S.header);

            // --- Merges: B/C/D vertical ---
            addMergeSafe(sh, new CellRangeAddress(hrTop.getRowNum(), hrBot.getRowNum(), colB, colB));
            addMergeSafe(sh, new CellRangeAddress(hrTop.getRowNum(), hrBot.getRowNum(), colC, colC));
            addMergeSafe(sh, new CellRangeAddress(hrTop.getRowNum(), hrBot.getRowNum(), colD, colD));

            // --- Merge group headers across two columns (top row) ---
            addMergeSafe(sh, new CellRangeAddress(hrTop.getRowNum(), hrTop.getRowNum(), colE, colF)); // 1인 제공량(g)
            addMergeSafe(sh, new CellRangeAddress(hrTop.getRowNum(), hrTop.getRowNum(), colG, colH)); // 총 발주량

            // --- Method: merge across columns and rows ---
            addMergeSafe(sh, new CellRangeAddress(hrTop.getRowNum(), hrTop.getRowNum(), methodFirst, methodLast));
            addMergeSafe(sh, new CellRangeAddress(hrBot.getRowNum(), hrBot.getRowNum(), methodFirst, methodLast));
            addMergeSafe(sh, new CellRangeAddress(hrTop.getRowNum(), hrBot.getRowNum(), methodFirst, methodLast));
        }

        // Split sections using morningCount

        // 오전간식
        if (!am.isEmpty()) {
            int startRowBlock = r;
            for (int i = 0; i < am.size(); i++) {
                String label = "";
                if (i == 0) label = "오전간식";
                r = writeMenuBlock(sh, S, tplNew, tplOld, T, kind, r, label, am.get(i));
            }
            mergeLabel(sh, startRowBlock, r - 1, 0 + OUT_COL_OFFSET);
        }

        // 점심
        if (!lunch.isEmpty()) {
            int startRowBlock = r;
            for (int i = 0; i < lunch.size(); i++) {
                String label = "";
                if (i == 0) label = "점심";
                r = writeMenuBlock(sh, S, tplNew, tplOld, T, kind, r, label, lunch.get(i));
            }
            mergeLabel(sh, startRowBlock, r - 1, 0 + OUT_COL_OFFSET);
        }

        // 오후간식
        if (!pm.isEmpty()) {
            int startRowBlock = r;
            for (int i = 0; i < pm.size(); i++) {
                String label = "";
                if (i == 0) label = "오후간식";
                r = writeMenuBlock(sh, S, tplNew, tplOld, T, kind, r, label, pm.get(i));

            }
            mergeLabel(sh, startRowBlock, r - 1, 0 + OUT_COL_OFFSET);
        }

        int lastDataRow = r - 1;
        int titleRowIndex = tr.getRowNum();
        if (lastDataRow >= titleRowIndex) {
            CellRangeAddress outer = new CellRangeAddress(titleRowIndex, lastDataRow, 0, LAST_COL);
            setMergedBorder(sh, outer, BorderStyle.DOUBLE);

            forceLeftOuterBorderDoubleAll(sh, titleRowIndex, lastDataRow, 1);
        }

        clearBordersInColumnA(sh, startRow, r - 1);

        r = appendAllergyImageUnderDay(sh, r, OUT_COL_OFFSET, LAST_COL, kind);

        return r;
    }

    /** Write one menu block (ingredients rows + method merge). */
    int writeMenuBlock(Sheet sh, Styles S, Workbook tplNew, Workbook tplOld, TemplateCols T, TplKind kind,
            int r, String label, String rawMenu) {

        int partStart = r;
        boolean noScaleP35 = isNoScaleMenu(rawMenu);

        Optional<Block> found = findBlockInTemplateExactPreferNew(tplNew, tplOld, rawMenu, T);

        if (found.isEmpty() || found.get().items.isEmpty()) {
            Row row = sh.createRow(r++);

            writeOneRow(kind, row, S, label, rawMenu,
                    "", "", "", "", "", "",
                    noScaleP35);
            setBordersRowBasic(kind, row, S);

            int methodFirstNF;
            if (kind == TplKind.P35) methodFirstNF = 7 + OUT_COL_OFFSET;
            else methodFirstNF = 5 + OUT_COL_OFFSET;

            int methodLastNF  = methodFirstNF + 4;
            addMergeSafe(sh, new CellRangeAddress(row.getRowNum(), row.getRowNum(), methodFirstNF, methodLastNF));
            for (int cc = methodFirstNF; cc <= methodLastNF; cc++) {
                Cell c = safeCell(sh, row.getRowNum(), cc);
                c.setCellStyle(S.methodMerged);
            }
            return r;
        }

        Block b = found.get();
        for (int i = 0; i < b.items.size(); i++) {
            Item it = b.items.get(i);
            Row row = sh.createRow(r++);
            String menuCell  = "";
            String methodTop = "";
            if (i == 0) {
                menuCell = b.displayName;
                methodTop = compactSpaces(b.method);
            }
            String lbl = "";
            if (i == 0) lbl = label;

            writeOneRow(kind, row, S, lbl,
                    menuCell,
                    nz(it.ingredient),
                    nz(it.p12), nz(it.p35), nz(it.t12), nz(it.t35),
                    methodTop,
                    noScaleP35);
        }

        int partEnd = r - 1;

        // Merge B (menu)
        if (partEnd > partStart) {
            addMergeSafe(sh, new CellRangeAddress(partStart, partEnd, 1 + OUT_COL_OFFSET, 1 + OUT_COL_OFFSET));
        }

        // Merge method area
        int methodFirst;
        if (kind == TplKind.P35) methodFirst = 7 + OUT_COL_OFFSET;
        else methodFirst = 5 + OUT_COL_OFFSET;

        int methodLast  = methodFirst + 4;
        addMergeSafe(sh, new CellRangeAddress(partStart, partEnd, methodFirst, methodLast));

        for (int rr = partStart; rr <= partEnd; rr++) {
            for (int cc = methodFirst; cc <= methodLast; cc++) {
                Cell c = safeCell(sh, rr, cc);
                c.setCellStyle(S.methodMerged);
            }
        }
        MethodCellLayout.fit(sh, new CellRangeAddress(partStart, partEnd, methodFirst, methodLast), rawMenu);
        return r;
    }

    /** Merge label in column col for range [start..end]. */
    void mergeLabel(Sheet sh, int start, int end, int col){
        if (end <= start) return;
        addMergeSafe(sh, new CellRangeAddress(start, end, col, col));
    }

    /** Safe merge: remove overlapping merged regions first, then add. */
    void addMergeSafe(Sheet sh, CellRangeAddress target){
        List<Integer> toRemove = new ArrayList<>();
        for (int i = 0; i < sh.getNumMergedRegions(); i++) {
            CellRangeAddress ex = sh.getMergedRegion(i);
            if (rangesOverlap(ex, target)) toRemove.add(i);
        }
        Collections.reverse(toRemove);
        for (int idx : toRemove) sh.removeMergedRegion(idx);
        sh.addMergedRegion(target);
        setMergedBorder(sh, target, BorderStyle.THIN);
    }

    boolean rangesOverlap(CellRangeAddress a, CellRangeAddress b){
        boolean rowOverlap = a.getFirstRow() <= b.getLastRow() && b.getFirstRow() <= a.getLastRow();
        boolean colOverlap = a.getFirstColumn() <= b.getLastColumn() && b.getFirstColumn() <= a.getLastColumn();
        return rowOverlap && colOverlap;
    }

    /** Write one row with per-template formulas. */
    void writeOneRow(TplKind kind, Row row, Styles S,
                            String label, String menu, String ing,
                            String p12, String p35, String t12, String t35,
                            String method,
                            boolean noScaleP35) {
        row.setHeightInPoints(18);

        Cell blankA = row.createCell(0);
        blankA.setCellValue("");
        blankA.setCellStyle(S.bodyCenter);

        Cell a = row.createCell(0 + OUT_COL_OFFSET);
        a.setCellValue(nz(label));
        CellStyle labelStyle = S.bodyCenter;
        String l = nz(label);
        if (l.equals("오전간식")) labelStyle = S.labelAm;
        else if (l.equals("점심")) labelStyle = S.labelLunch;
        else if (l.equals("오후간식")) labelStyle = S.labelPm;
        a.setCellStyle(labelStyle);

        Cell b = row.createCell(1 + OUT_COL_OFFSET); b.setCellValue(nz(menu));  b.setCellStyle(S.bodyCenter);
        Cell c = row.createCell(2 + OUT_COL_OFFSET); c.setCellValue(nz(ing));   c.setCellStyle(S.bodyCenter);

        if (kind == TplKind.P35) {
            Cell d = row.createCell(3 + OUT_COL_OFFSET);
            int excelRow = row.getRowNum() + 1;

            if (noScaleP35) {
                // English comment: No 0.65 scaling for specific menus (milk/soy/cheese/yogurt).
                // Use p35 value directly (same as 3~5 serving amount).
                Double p35v = parseNumericOrNull(p35);
                if (p35v != null) {
                    double dVal = p35v;
                    if (isIntegerDouble(dVal)) d.setCellStyle(S.num1);
                    else {
                        double one = Math.round(dVal * 10.0) / 10.0;
                        if (Math.abs(dVal - one) > 1e-9) d.setCellStyle(S.dec1);
                        else d.setCellStyle(S.num1);
                    }
                    // Write numeric
                    double rounded = Math.round(dVal * 10.0) / 10.0;
                    if (Math.abs(rounded - Math.rint(rounded)) < 1e-9) d.setCellValue((long) Math.rint(rounded));
                    else d.setCellValue(rounded);
                } else {
                    // If p35 is not numeric, keep blank (avoid #VALUE!).
                    d.setCellStyle(S.num1);
                }
            } else {
                // Existing behavior: 1~2 = 0.65 * (3~5)
                d.setCellFormula(String.format("F%d*0.65", excelRow));

                Double p35v = parseNumericOrNull(p35);
                if (p35v != null) {
                    double dVal = p35v * 0.65;
                    if (isIntegerDouble(dVal)) {
                        d.setCellStyle(S.num1);
                    } else {
                        double one = Math.round(dVal * 10.0) / 10.0;
                        if (Math.abs(dVal - one) > 1e-9) d.setCellStyle(S.dec1);
                        else d.setCellStyle(S.num1);
                    }
                } else {
                    d.setCellStyle(S.num1);
                }
            }

            Double p35v = parseNumericOrNull(p35);
            if (p35v != null) {
                double dVal = p35v * 0.65;
                if (isIntegerDouble(dVal)) {
                    d.setCellStyle(S.num1);
                } else {
                    double one = Math.round(dVal * 10.0) / 10.0;
                    if (Math.abs(dVal - one) > 1e-9) d.setCellStyle(S.dec1);
                    else d.setCellStyle(S.num1);
                }
            } else {
                d.setCellStyle(S.num1);
            }

            Cell e = row.createCell(4 + OUT_COL_OFFSET);
            // English comment: Do NOT write empty string into numeric cells; keep BLANK to avoid #VALUE! in formulas
            if (!trySetNumeric(e, p35)) {
                if (!isBlank(p35)) e.setCellValue(nz(p35));
            }
            e.setCellStyle(S.num1);

            Cell t_12 = row.createCell(5 + OUT_COL_OFFSET);
            t_12.setCellFormula(String.format("E%d*$E$2", excelRow));
            t_12.setCellStyle(S.num1);

            Cell t_35 = row.createCell(6 + OUT_COL_OFFSET);
            t_35.setCellFormula(String.format("F%d*$E$3", excelRow));
            t_35.setCellStyle(S.num1);

            Cell g = row.createCell(7 + OUT_COL_OFFSET);
            g.setCellValue(compactSpaces(nz(method)));

            setBordersRowBasic(kind, row, S);

        } else {
            Cell d = row.createCell(3 + OUT_COL_OFFSET);
            // English comment: Do NOT write empty string into numeric cells; keep BLANK to avoid #VALUE! in formulas
            if (!trySetNumeric(d, p12)) {
                if (!isBlank(p12)) d.setCellValue(nz(p12));
            }
            d.setCellStyle(S.num1);

            int excelRow = row.getRowNum() + 1;
            Cell e = row.createCell(4 + OUT_COL_OFFSET);
            e.setCellFormula(String.format("E%d*$E$2", excelRow));
            e.setCellStyle(S.num1);

            Cell f = row.createCell(5 + OUT_COL_OFFSET);
            f.setCellValue(compactSpaces(nz(method)));

            setBordersRowBasic(kind, row, S);
        }
    }

    void setBordersRowBasic(TplKind kind, Row row, Styles S){
        int lastCol;
        if (kind == TplKind.P35) lastCol = 6;
        else lastCol = 4;

        for (int cc = 0; cc <= lastCol; cc++) setBorders(safeCell(row.getSheet(), row.getRowNum(), cc), BorderStyle.THIN);
    }

    /* ===== Template search: exact only ===== */
    Optional<Block> findBlockInTemplateExact(Workbook wb, String menuRaw, TemplateCols T){
        String keyNorm = normalizeForMatch(menuRaw);
        for (int s = wb.getNumberOfSheets() - 1; s >= 0; s--) {
            Sheet sh = wb.getSheetAt(s);
            int lastRow = sh.getLastRowNum();
            for (int r = 0; r <= lastRow; r++) {
                String cellMenu = readStringConsideringMerged(sh, r, T.colMenu);
                if (isBlank(cellMenu)) continue;
                String cellNorm = normalizeForMatch(cellMenu);
                if (!cellNorm.isEmpty() && cellNorm.equals(keyNorm)) {
                    return Optional.of(buildBlockFromAnchor(sh, r, T));
                }
            }
        }
        return Optional.empty();
    }

    Optional<Block> findBlockInTemplateExactPreferNew(
            Workbook tplNew,
            Workbook tplOld,
            String menuRaw,
            TemplateCols T
    ) {
        // English comment: Try 2026 template first, then fallback to 2022~2025
        if (tplNew != null) {
            Optional<Block> a = findBlockInTemplateExact(tplNew, menuRaw, T);
            if (a.isPresent()) return a;
        }
        if (tplOld != null) {
            Optional<Block> b = findBlockInTemplateExact(tplOld, menuRaw, T);
            if (b.isPresent()) return b;
        }
        return Optional.empty();
    }

    /** Build block starting from menu anchor row. */
    Block buildBlockFromAnchor(Sheet sh, int r, TemplateCols T){
        int lastRow = sh.getLastRowNum();
        String cellMenu = readStringConsideringMerged(sh, r, T.colMenu);

        CellRangeAddress menuRange = findMergedRange(sh, r, T.colMenu);
        int first = r;
        int last = r;

        if (menuRange != null) {
            first = menuRange.getFirstRow();
            last = menuRange.getLastRow();
        } else {
            int rr = r + 1;
            while (rr <= lastRow) {
                String nextMenu = readStringConsideringMerged(sh, rr, T.colMenu);
                if (!isBlank(nextMenu)) break;
                rr++;
            }
            last = rr - 1;
        }

        String methodTop = "";
        outer:
        for (int rr = first; rr <= last; rr++) {
            for (int cc = T.colMethodStart; cc <= T.colMethodEndIncl; cc++) {
                String v = readStringConsideringMerged(sh, rr, cc);
                if (!isBlank(v)) { methodTop = v.trim(); break outer; }
            }
        }

        List<Item> items = new ArrayList<>();
        for (int rr = first; rr <= last; rr++) {
            String ing = readStringConsideringMerged(sh, rr, T.colIng);
            String c12 = "";
            String c35 = "";
            String t12 = "";
            String t35 = "";

            if (T.colP12 >= 0) c12 = readStringConsideringMerged(sh, rr, T.colP12);
            if (T.colP35 >= 0) c35 = readStringConsideringMerged(sh, rr, T.colP35);
            if (T.colT12 >= 0) t12 = readStringConsideringMerged(sh, rr, T.colT12);
            if (T.colT35 >= 0) t35 = readStringConsideringMerged(sh, rr, T.colT35);

            boolean rowEmpty = isBlank(ing) && isBlank(c12) && isBlank(c35) && isBlank(t12) && isBlank(t35);
            if (rowEmpty) continue;
            if ("식재료명".equals(ing)) continue;

            Item it = new Item();
            it.ingredient = ing;
            it.p12 = c12; it.p35 = c35; it.t12 = t12; it.t35 = t35;
            items.add(it);
        }

        Block b = new Block();
        b.displayName = cellMenu;
        b.method = nz(methodTop);
        b.items  = items;
        return b;
    }

    /* ===== Utils ===== */
    List<String> compressConsecutive(List<String> src){
        if (src == null || src.isEmpty()) return new ArrayList<>();
        List<String> out = new ArrayList<>();
        String prev = null;
        for (String s : src) {
            String t = nz(s).trim();
            if (t.isEmpty()) continue;
            if (t.equals("없음")) continue;
            if (prev != null && normalizeMenu(prev).equals(normalizeMenu(t))) {
                // skip duplicate
            } else {
                out.add(t);
                prev = t;
            }
        }
        return out;
    }

    boolean isKimchi(String menu){
        String k = normalizeMenu(menu);
        if (k.contains("김치")) return true;
        if (k.contains("깍두기")) return true;
        return false;
    }

    Path deriveOutXlsxPathFromJson(Path jsonPath, Path outDir){
        String file = jsonPath.getFileName().toString();
        String stem = file.contains(".") ? file.substring(0, file.lastIndexOf('.')) : file;
        String outBase = stem + "_수정";

        // English comment: Keep pattern but avoid optional syntax that causes formula parsing issues
        Pattern p = Pattern.compile("^(\\d{2})\\.(\\d{1,2})\\.\\s*만(\\d+(-\\d+){0,1})세\\s*(일반형|시간연장형)$");
        if (p.matcher(stem).find()) outBase = stem + "_수정";
        return outDir.resolve(outBase + ".xlsx");
    }

    YearMonth inferYMFromJson(Path jsonPath) {
        String name = jsonPath.getFileName().toString();

        Pattern pYY = Pattern.compile("^(\\d{2})\\.(\\d{1,2})");
        Matcher mYY = pYY.matcher(name);
        if (mYY.find()) {
            int yy = Integer.parseInt(mYY.group(1));
            int mm = Integer.parseInt(mYY.group(2));
            return YearMonth.of(2000 + yy, Math.min(Math.max(mm,1),12));
        }

        Pattern pYYYY = Pattern.compile("(20\\d{2})[\\.-_/ ]{0,1}(\\d{1,2})");
        Matcher mYYYY = pYYYY.matcher(name);
        if (mYYYY.find()) {
            int y = Integer.parseInt(mYYYY.group(1));
            int mo = Integer.parseInt(mYYYY.group(2));
            return YearMonth.of(y, Math.min(Math.max(mo,1),12));
        }

        YearMonth ym2 = tryInferYMFromKoreanName(name);
        if (ym2 != null) return ym2;

// English comment: Last resort: use current year-month.
        return YearMonth.from(LocalDate.now());
    }

    List<DayPlan> readPlan(Path json) throws IOException {
        ObjectMapper om = new ObjectMapper();
        try (var is = Files.newInputStream(json)) {
            return om.readValue(is, new TypeReference<List<DayPlan>>() {});
        }
    }


    // English comment: Read plan list from JSON string (no intermediate JSON file)
    List<DayPlan> readPlanFromString(String jsonString) throws IOException {
        if (jsonString == null) throw new IllegalArgumentException("jsonString is null");
        ObjectMapper om = new ObjectMapper();
        return om.readValue(jsonString, new TypeReference<List<DayPlan>>() {
        });
    }

    private Path doConvertFromJsonString(String jsonString, Path sourceNamePath, Path outputXlsx) throws Exception {
        if (jsonString == null) throw new IllegalArgumentException("jsonString is null");
        if (sourceNamePath == null) throw new IllegalArgumentException("sourceNamePath is null");
        if (outputXlsx == null) throw new IllegalArgumentException("outputXlsx is null");
        return render(jsonString, sourceNamePath, outputXlsx, inferYMFromSourceName(sourceNamePath));
    }

    private Path render(String jsonString, Path sourceNamePath, Path outputXlsx, YearMonth suppliedMonth) throws Exception {
        BORDER_STYLE_CACHE.clear();

        ZipSecureFile.setMinInflateRatio(0.0d);
        ZipSecureFile.setMaxFileCount(20000);

        // English comment: Validate inputs
        if (jsonString == null) throw new IllegalArgumentException("jsonString is null");
        if (sourceNamePath == null) throw new IllegalArgumentException("sourceNamePath is null");

        if (outputXlsx == null) {
            throw new IllegalArgumentException("outputXlsx is null");
        }

        YearMonth ym = suppliedMonth;
        String titlePrefix = String.format("%d년 %d월", ym.getYear(), ym.getMonthValue());

        String monthPrefix = extractMonthPrefixFromSourceName(sourceNamePath); // English comment: "26.1월"

        Path tplFile = resolveGeneralTemplate(sourceNamePath);
        Path tplNewPath = resolveGeneralTemplateNew(sourceNamePath);
        Path tplOldPath = resolveGeneralTemplateOld(sourceNamePath);

        TplKind tplKind = detectTplKind(tplFile);
        TemplateCols T = colsOf(tplKind);

        List<DayPlan> plan = readPlanFromString(jsonString);

        if (outputXlsx.getParent() != null) Files.createDirectories(outputXlsx.getParent());

        try (Workbook out = WorkbookResources.open(tplFile);
             Workbook tplNew = WorkbookResources.open(tplNewPath);
             Workbook tplOld = WorkbookResources.open(tplOldPath)) {


            // English comment: Split normal days and birthday days
            List<DayPlan> normalDays = new ArrayList<>();
            List<DayPlan> birthdayDays = new ArrayList<>();

            for (DayPlan d : plan) {
                if (d == null) continue;
                String ds = nz(d.dateStr).trim();
                if (ds.equalsIgnoreCase("birthday")) birthdayDays.add(d);
                else normalDays.add(d);
            }

            int weekIndex = 1;

            String baseName = out.getSheetName(0);

            // Keep only base sheet in template workbook
            for (int i = out.getNumberOfSheets() - 1; i >= 0; i--) {
                String nm = out.getSheetName(i);
                if (!baseName.equals(nm)) out.removeSheetAt(i);
            }

            // Re-find baseIdx (it will be 0 after pruning)
            int baseIdx = out.getSheetIndex(baseName);
            out.setSheetName(baseIdx, "__BASE_TEMPLATE__");
            Styles S = Styles.build(out);

            int DATA_START_ROW;
            if (tplKind == TplKind.P12) DATA_START_ROW = 5;  // 1-2세: title at Excel row 6
            else DATA_START_ROW = 6;                           // 3-5세: title at Excel row 7

            int keepLastRow = DATA_START_ROW - 1;

            Sheet sh = copyTopTemplateArea(
                    out,
                    baseIdx,
                    buildSheetName(monthPrefix, weekIndex),
                    keepLastRow
            );

            int currentRow = DATA_START_ROW;


            for (int i = 0; i < normalDays.size(); i++) {
                DayPlan d = normalDays.get(i);
                if (d == null) continue;
                if (d.date == null) continue;

                int beforeRow = currentRow;

                currentRow = writeOneDay(sh, S, d, tplNew, tplOld, T, tplKind, titlePrefix, currentRow, false);

                if (currentRow > beforeRow) {
                    currentRow = currentRow + 1;
                }

                if (isWeekEndGeneral(d.weekday)) {
                    if (i < normalDays.size() - 1) {
                        weekIndex = weekIndex + 1;
                        sh = copyTopTemplateArea(
                                out,
                                baseIdx,
                                buildSheetName(monthPrefix, weekIndex),
                                keepLastRow
                        );

                        currentRow = DATA_START_ROW;
                    }
                }
            }

            // English comment: Birthday dedicated sheet
            if (!birthdayDays.isEmpty()) {
                keepLastRow = DATA_START_ROW - 1;

                Sheet bdaySheet = copyTopTemplateArea(
                        out,
                        baseIdx,
                        buildBirthdaySheetName(monthPrefix),
                        keepLastRow
                );

                int r = DATA_START_ROW;

                for (DayPlan d : birthdayDays) {
                    r = writeOneDay(bdaySheet, S, d, tplNew, tplOld, T, tplKind, titlePrefix, r, true);
                    r = r + 4; // Birthday는 하루 블록 간격을 넉넉히
                }
            }

            int baseIdx2 = out.getSheetIndex("__BASE_TEMPLATE__");
            if (baseIdx2 >= 0) out.removeSheetAt(baseIdx2);

            try (OutputStream os = Files.newOutputStream(outputXlsx)) {
                out.write(os);
            }
        }

        return outputXlsx;
        }


    static String normalizeForMatch(String s){
        if (s == null) return "";
        String t = MenuNamePrefix.withoutStep(s);
        t = t.replaceAll("[\u2460-\u2472]", "");
        t = t.replace("쇠고기", "소고기").replace("닭 살", "닭살").replace("계란", "달걀");
        t = t.replaceAll("\\s+","")
                .replaceAll("[()\\[\\]{}／/·ㆍ・＆&\\-]", "")
                .replaceAll("[★☆※•…]", "");
        return t;
    }

    String normalizeMenu(String s) { return normalizeForMatch(s); }

    CellRangeAddress findMergedRange(Sheet sh, int r, int c){
        for (int i = 0; i < sh.getNumMergedRegions(); i++) {
            CellRangeAddress ra = sh.getMergedRegion(i);
            if (ra.isInRange(r, c)) return ra;
        }
        return null;
    }

    String readStringConsideringMerged(Sheet sh, int r, int c){
        Cell cell = getMergedAnchorCell(sh, r, c);
        if (cell == null) return "";
        return getString(cell).trim();
    }

    Cell getMergedAnchorCell(Sheet sh, int r, int c){
        for (int i=0;i<sh.getNumMergedRegions();i++){
            CellRangeAddress ra = sh.getMergedRegion(i);
            if (ra.isInRange(r,c)){
                Row topRow = sh.getRow(ra.getFirstRow());
                if (topRow==null) return null;
                return topRow.getCell(ra.getFirstColumn());
            }
        }
        Row row = sh.getRow(r);
        if (row==null) return null;
        return row.getCell(c);
    }

    String getString(Cell cell){
        if (cell==null) return "";
        switch (cell.getCellType()){
            case STRING: return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) return cell.getDateCellValue().toString();
                double v = cell.getNumericCellValue();
                if (Math.abs(v - Math.round(v)) < 1e-9) return String.valueOf((long)Math.round(v));
                return String.valueOf(v);
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try { return cell.getStringCellValue(); }
                catch (Exception e){
                    try { return String.valueOf(cell.getNumericCellValue()); }
                    catch (Exception ignore){ return ""; }
                }
            default: return "";
        }
    }

    boolean trySetNumeric(Cell c, String s) {
        if (isBlank(s)) return false;
        try {
            String t = s.replace(",", "")
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
            double rounded = Math.round(v * 10.0) / 10.0;
            if (Math.abs(rounded - Math.rint(rounded)) < 1e-9) c.setCellValue((long) Math.rint(rounded));
            else c.setCellValue(rounded);
            return true;
        } catch (Exception e) { return false; }
    }

    Double parseNumericOrNull(String s) {
        if (isBlank(s)) return null;
        try {
            String t = s.replace(",", "")
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
        } catch (Exception e) { return null; }
    }

    boolean isIntegerDouble(double v){
        return Math.abs(v - Math.rint(v)) < 1e-9;
    }

    void setBorders(Cell cell, BorderStyle bs) {
        if (cell == null) return;

        Workbook wb = cell.getSheet().getWorkbook();
        CellStyle base = cell.getCellStyle();

        // English comment: Key uses base style index and border style name.
        String key = base.getIndex() + "|" + bs.name();

        CellStyle cached = BORDER_STYLE_CACHE.get(key);
        if (cached == null) {
            CellStyle clone = wb.createCellStyle();
            clone.cloneStyleFrom(base);
            clone.setBorderBottom(bs);
            clone.setBorderTop(bs);
            clone.setBorderLeft(bs);
            clone.setBorderRight(bs);
            BORDER_STYLE_CACHE.put(key, clone);
            cached = clone;
        }

        cell.setCellStyle(cached);
    }

    void setMergedBorder(Sheet sh, CellRangeAddress rgn, BorderStyle bs) {
        org.apache.poi.ss.util.RegionUtil.setBorderTop(bs, rgn, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderBottom(bs, rgn, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderLeft(bs, rgn, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderRight(bs, rgn, sh);
    }

    Cell safeCell(Sheet sh, int r, int c) {
        Row row = sh.getRow(r);
        if (row == null) row = sh.createRow(r);
        Cell cell = row.getCell(c);
        if (cell == null) cell = row.createCell(c);
        return cell;
    }

    boolean isBlank(String s){ return s==null || s.trim().isEmpty(); }
    String nz(String s){ return s==null ? "" : s; }

    String compactSpaces(String s){
        if (s == null) return "";
        return s.replace('\u00A0',' ')
                .replaceAll("[ \\t]{2,}", " ")
                .trim();
    }

    void clearBordersInColumnA(Sheet sh, int fromRow, int toRow) {
        for (int r = fromRow; r <= toRow; r++) {
            Row row = sh.getRow(r);
            if (row == null) continue;

            Cell cell = row.getCell(0);
            if (cell == null) continue;

            CellStyle cs = sh.getWorkbook().createCellStyle();
            cs.cloneStyleFrom(cell.getCellStyle());
            cs.setBorderTop(BorderStyle.NONE);
            cs.setBorderBottom(BorderStyle.NONE);
            cs.setBorderLeft(BorderStyle.NONE);
            cs.setBorderRight(BorderStyle.NONE);
            cell.setCellStyle(cs);
        }
    }

    Sheet copyTopTemplateArea(
            Workbook out,
            int baseIdx,
            String newSheetName,
            int keepLastRow
    ) {
        // 1) Clone sheet (copies column widths, row heights, merged regions, images, drawings)
        Sheet newSh = out.cloneSheet(baseIdx);

        // 2) Rename
        int newIdx = out.getSheetIndex(newSh);
        out.setSheetName(newIdx, newSheetName);

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

        newSh.setFitToPage(true);
        PrintSetup ps = newSh.getPrintSetup();
        ps.setFitWidth((short) 1);
        ps.setFitHeight((short) 0);

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

    Path resolveAllergyImagePath() {
        Path appHome = getAppHomeDir();
        Path p1 = appHome.resolve("input").resolve("allergy.png");
        if (Files.exists(p1)) return p1;

        Path p2 = Paths.get("").toAbsolutePath().resolve("input").resolve("allergy.png");
        if (Files.exists(p2)) return p2;

        return null;
    }

    // English comment: Append allergy image under a day block, return next row index
    int appendAllergyImageUnderDay(Sheet sh, int rowIndex, int firstCol, int lastCol, TplKind kind) {
        try {
            Path imgPath = resolveAllergyImagePath();
            if (imgPath == null) return rowIndex;

            byte[] imgBytes = Files.readAllBytes(imgPath);
            int picIdx = sh.getWorkbook().addPicture(imgBytes, Workbook.PICTURE_TYPE_PNG);

            // English comment: Create a single "image row" with enough height
            Row imgRow = sh.getRow(rowIndex);
            if (imgRow == null) imgRow = sh.createRow(rowIndex);

            // English comment: Adjust height if you want bigger/smaller image under each day
            imgRow.setHeightInPoints(80.0f);

            Drawing<?> drawing = sh.createDrawingPatriarch();
            CreationHelper helper = sh.getWorkbook().getCreationHelper();

            ClientAnchor anchor = helper.createClientAnchor();
            anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);

            // English comment: Span the same width as the day table (B..LAST_COL)
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

            // English comment: Fit image into the anchor cell range
            pic.resize(1.00);

            // English comment: Clear borders in column A for the image row too (keep layout clean)
            clearBordersInColumnA(sh, rowIndex, rowIndex);

            return rowIndex + 1;

        } catch (Exception ignore) {
            // English comment: Ignore image errors to avoid breaking XLSX generation
            return rowIndex;
        }
    }

    void forceLeftOuterBorderDoubleAll(Sheet sh, int firstRow, int lastRow, int leftCol) {

        // English comment: 1) Force cell styles (for non-merged cells).
        int rr = firstRow;
        while (rr <= lastRow) {
            Cell c = safeCell(sh, rr, leftCol);

            CellStyle cs = sh.getWorkbook().createCellStyle();
            cs.cloneStyleFrom(c.getCellStyle());

            cs.setBorderLeft(BorderStyle.DOUBLE);

            c.setCellStyle(cs);
            rr = rr + 1;
        }

        // English comment: 2) Force merged regions that sit on the left edge too (RegionUtil wins in Excel).
        int i = 0;
        while (i < sh.getNumMergedRegions()) {
            CellRangeAddress ra = sh.getMergedRegion(i);

            boolean touchesLeftEdge = (ra.getFirstColumn() == leftCol);
            boolean overlapsRows = !(ra.getLastRow() < firstRow || ra.getFirstRow() > lastRow);

            if (touchesLeftEdge && overlapsRows) {
                org.apache.poi.ss.util.RegionUtil.setBorderLeft(BorderStyle.DOUBLE, ra, sh);
            }

            i = i + 1;
        }
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

        if (sheetName.length() > 31) sheetName = sheetName.substring(0, 31);

        sheetName = sheetName.replace("/", "_")
                .replace("\\", "_")
                .replace("[", "(")
                .replace("]", ")")
                .replace(":", "-")
                .replace("*", "_");

        return sheetName;
    }


    private String buildBirthdaySheetName(String monthPrefix) {
        // English comment: Build "26.1월 ★birthday"
        String sheetName = monthPrefix.replace(".", ". ") + " ★birthday";

        if (sheetName.length() > 31) {
            sheetName = sheetName.substring(0, 31);
        }

        sheetName = sheetName.replace("/", "_");
        sheetName = sheetName.replace("\\", "_");
        sheetName = sheetName.replace("[", "(");
        sheetName = sheetName.replace("]", ")");
        sheetName = sheetName.replace(":", "-");
        sheetName = sheetName.replace("*", "_");

        return sheetName;
    }

    YearMonth inferYMFromSourceName(Path sourceNamePath) {
        // English comment: Parse "2026년 1월 ..." from HWP file name.
        String name = sourceNamePath.getFileName().toString();

        int dot = name.lastIndexOf('.');
        if (dot > 0) {
            name = name.substring(0, dot);
        }

        Pattern p = Pattern.compile("(\\d{4})\\s*년\\s*(\\d{1,2})\\s*월");
        Matcher m = p.matcher(name);

        if (m.find()) {
            int y = Integer.parseInt(m.group(1));
            int mo = Integer.parseInt(m.group(2));

            if (mo < 1) mo = 1;
            if (mo > 12) mo = 12;

            return YearMonth.of(y, mo);
        }

        // English comment: Fallback to previous parser (may work for other naming patterns).
        return inferYMFromJson(sourceNamePath);
    }

    YearMonth tryInferYMFromKoreanName(String filename) {
        if (filename == null) return null;

        String name = filename;

        int dot = name.lastIndexOf('.');
        if (dot > 0) name = name.substring(0, dot);

        Matcher m = Pattern.compile("(20\\d{2})\\s*년\\s*(\\d{1,2})\\s*월").matcher(name);
        if (!m.find()) return null;

        int y = Integer.parseInt(m.group(1));
        int mo = Integer.parseInt(m.group(2));

        if (mo < 1) mo = 1;
        if (mo > 12) mo = 12;

        return YearMonth.of(y, mo);
    }

    boolean isNoScaleMenu(String menuRaw) {
        if (menuRaw == null) return false;
        String n = normalizeForMatch(menuRaw);
        return NO_SCALE_MENUS_NORM.contains(n);
    }

    // New: Convert JSON -> EXCEL with explicit paths (for AllInOne)
    private Path doConvert(Path inputJson, Path outputXlsx) throws Exception {
        return render(Files.readString(inputJson, java.nio.charset.StandardCharsets.UTF_8), inputJson, outputXlsx, inferYMFromJson(inputJson));
    }

    public static Path convertFromJsonString(String json, Path source, Path output, Path appHome) throws Exception {
        return new JsonToExcelGeneral(appHome).doConvertFromJsonString(json, source, output);
    }

    public static void main(String[] args) throws Exception {
        new JsonToExcelGeneral().runMain(args);
    }

    public static Path convertFromJsonString(String jsonString, Path sourceNamePath, Path outputXlsx) throws Exception {
        return new JsonToExcelGeneral().doConvertFromJsonString(jsonString, sourceNamePath, outputXlsx);
    }

    public static Path convert(Path inputJson, Path outputXlsx) throws Exception {
        return new JsonToExcelGeneral().doConvert(inputJson, outputXlsx);
    }

}
