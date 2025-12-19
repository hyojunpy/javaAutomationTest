// (전체 파일) org/example/HwpToJson.java
package org.example;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

// ====== hwplib imports ======
import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.object.bodytext.Section;
import kr.dogfoot.hwplib.object.bodytext.ParagraphListInterface;
import kr.dogfoot.hwplib.object.bodytext.control.Control;
import kr.dogfoot.hwplib.object.bodytext.control.ControlTable;
import kr.dogfoot.hwplib.object.bodytext.control.ControlType;
import kr.dogfoot.hwplib.object.bodytext.control.gso.GsoControl;
import kr.dogfoot.hwplib.object.bodytext.control.gso.caption.Caption;
import kr.dogfoot.hwplib.object.bodytext.control.table.Cell;
import kr.dogfoot.hwplib.object.bodytext.control.table.Row;
import kr.dogfoot.hwplib.object.bodytext.paragraph.Paragraph;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.HWPChar;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.HWPCharNormal;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.HWPCharType;
import kr.dogfoot.hwplib.reader.HWPReader;
// =======================================

import java.io.IOException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.io.File;

public class HwpToJson {
    // ===== Input (default) =====
    private static final Path IN_HWP = Paths.get("input", "2026년 1월 시간연장형(만1-2세).hwp");

    // ===== Output dir =====
    private static final Path OUT_DIR = Paths.get("output");

    private static HWPFile HWP;

    // “1(월)” like header (tolerate trailing text)
    private static final Pattern DAY_CELL = Pattern.compile("^(\\d{1,2})\\((.)\\)");

    // collapse multi-space characters
    private static final Pattern SPACES_ONLY = Pattern.compile("[ \\t\\u00A0\\u2000-\\u200B\\u202F\\u205F\\u3000]+");

    // drop tokens / lines
    private static final Set<String> DROP_TOKENS_MISC = Set.of("열량/단백질", "일자/요일", "없음");
    private static final Set<String> DROP_TOKENS_FOODS = Set.of();
    private static final List<String> DROP_LINE_HINTS = List.of(
            "<식품 알레르기", "원 산 지 표 시", "수산물", "가공품", "국가명(", "[1월 제철 식재료]", "[12월 제철 식재료]"
    );

    // allergens handling
    private static final Pattern ALLERGENS = Pattern.compile("^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲]+$");

    // parsed year-month from filename
    private static Integer PARSED_YEAR = null;
    private static Integer PARSED_MONTH = null;

    // ===== Model =====
    public static class DayPlan {
        public Integer date;
        public String  weekday;
        public List<String> menus = new ArrayList<>();
        DayPlan(int d, String w){ date = d; weekday = w; }
    }

    public static void main(String[] args) throws Exception {
        Path in = (args != null && args.length > 0 && args[0] != null && !args[0].isBlank())
                ? Paths.get(args[0])
                : IN_HWP;
        Files.createDirectories(OUT_DIR);

        // derive pretty output basename like "25.01. 만3-5세 시간연장형"
        String prettyBase = derivePrettyBasename(in.getFileName().toString());
        Path outJson = OUT_DIR.resolve(prettyBase + ".json");
        Path outTsv  = OUT_DIR.resolve(prettyBase + "-debug.tsv");

        int[] ym = parseYearMonthFromFilename(in.getFileName().toString());
        if (ym != null) { PARSED_YEAR = ym[0]; PARSED_MONTH = ym[1]; }

        String inAbs = in.toAbsolutePath().toString();

        System.out.println("[INPUT ] " + in.toAbsolutePath());
        System.out.println("[OUTPUT] " + outJson.toAbsolutePath());
        System.out.println("[DEBUG ] " + outTsv.toAbsolutePath());

        HWP = HWPReader.fromFile(inAbs);

        Map<Integer, DayPlan> byDate = new LinkedHashMap<>();

        // scan all sections/paragraphs; process ControlType.Table only
        for (int s = 0; s < HWP.getBodyText().getSectionList().size(); s++) {
            Section sec = HWP.getBodyText().getSectionList().get(s);
            for (int p = 0; p < sec.getParagraphCount(); p++) {
                Paragraph para = sec.getParagraph(p);
                if (para.getControlList() == null) continue;
                for (Control c : para.getControlList()) {
                    if (c.getType() != ControlType.Table) continue;
                    ControlTable ct = (ControlTable) c;
                    scanTable(ct, byDate, outTsv);
                }
            }
        }

        // finalize: ensure non-empty menus
        for (DayPlan dp : byDate.values()) {
            if (dp.menus == null || dp.menus.isEmpty()) dp.menus = List.of("없음");
        }

        List<DayPlan> list = new ArrayList<>(byDate.values());
        list.sort(Comparator.comparingInt(dp -> dp.date));

        new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT)
                .writeValue(outJson.toFile(), list);

        System.out.println("[OK] JSON written -> " + outJson.toAbsolutePath());
    }

    // ====== scan a table (dinner only) ======
    private static void scanTable(ControlTable ct, Map<Integer, DayPlan> byDate, Path outTsv) throws IOException {
        List<Row> rows = ct.getRowList();
        if (rows == null || rows.isEmpty()) return;


        List<String> currentHeader = null;
        List<Row> dinnerRows = new ArrayList<>();
        int r = 0;

        while (r < rows.size()) {
            String first = readCellFlat(rows.get(r), 0);

            // header start
            if (first.contains("일자/요일")) {
                if (currentHeader != null) flushDinnerBlock(currentHeader, dinnerRows, byDate);
                currentHeader = collectRowTextsFlat(rows.get(r));
                dinnerRows.clear();
                r++;
                continue;
            }

            // collect dinner block after header
            if (currentHeader != null) {
                if (isDinnerLabel(first)) {
                    int j = r;
                    dinnerRows.clear();
                    for (; j < rows.size(); j++) {
                        String f = readCellFlat(rows.get(j), 0);
                        if (f.contains("열량/단백질") || f.contains("일자/요일")) break;
                        dinnerRows.add(rows.get(j));
                    }
                    flushDinnerBlock(currentHeader, dinnerRows, byDate);

                    if (j < rows.size() && readCellFlat(rows.get(j), 0).contains("일자/요일")) {
                        currentHeader = collectRowTextsFlat(rows.get(j));
                        dinnerRows.clear();
                        r = j + 1;
                    } else {
                        r = j;
                    }
                    continue;
                }
            }
            r++;
        }

        if (currentHeader != null) flushDinnerBlock(currentHeader, dinnerRows, byDate);
    }

    // apply dinner rows to dates of header
    private static void flushDinnerBlock(List<String> header, List<Row> dinnerRows, Map<Integer, DayPlan> byDate) throws IOException {
        if (header == null || header.size() <= 1) return;

        int cStart = findFirstUsableHeaderIndex(header);
        if (cStart < 1) return;

        for (int c = cStart; c < header.size(); c++) {
            String rawHeader = header.get(c);
            String flat = stripSpacesFlat(rawHeader);

            int effC = effectiveMenuCol(c, header);

            Matcher m = DAY_CELL.matcher(flat);
            if (!m.find()) continue;

            int date = Integer.parseInt(m.group(1));
            String headerDow = m.group(2);

            String usedDow;
            if (PARSED_YEAR != null && PARSED_MONTH != null) {
                String calc = safeDowKorean(PARSED_YEAR, PARSED_MONTH, date);
                if (calc != null) {
                    if ("일".equals(calc)) continue;      // Skip Sundays
                    if (!calc.equals(headerDow)) continue;
                    usedDow = calc;
                } else {
                    usedDow = headerDow;
                }
            } else {
                usedDow = headerDow;
            }

            DayPlan dp = byDate.computeIfAbsent(date, d -> new DayPlan(d, usedDow));

            boolean holiday = isHolidayHeader(flat);

            // PROBE: find best body col around effC (handles merged cells shifting)
            int bestCol = probeBestBodyCol(dinnerRows, effC, 2);
            List<String> menus = collectDinnerMenus(dinnerRows, bestCol);

            if (holiday) {
                dp.menus = List.of("없음");
            } else {
                if (menus.isEmpty()) dp.menus = List.of("없음");     // safety
                else dp.menus = new ArrayList<>(menus);
            }
        }
    }

    // probe nearby columns to find any non-empty cell
    private static int probeBestBodyCol(List<Row> dinnerRows, int centerCol, int radius) throws IOException {
        int best = centerCol;
        if (hasAnyTextAtCol(dinnerRows, centerCol)) return centerCol;

        int min = Math.max(1, centerCol - radius);
        int max = centerCol + radius;

        int c = min;
        while (c <= max) {
            if (c != centerCol && hasAnyTextAtCol(dinnerRows, c)) return c;
            c = c + 1;
        }
        return best;
    }

    private static boolean hasAnyTextAtCol(List<Row> rows, int col) throws IOException {
        if (rows == null || rows.isEmpty()) return false;
        for (Row r : rows) {
            String label = readCellFlat(r, 0);
            if (label.contains("일자/요일") || label.contains("열량/단백질")) break;
            String v = readCellKeepNewlines(r, col);
            if (!stripSpacesFlat(v).isEmpty()) return true;
        }
        return false;
    }

    // collect all dinner menus for one column
    private static List<String> collectDinnerMenus(List<Row> dinnerRows, int col) throws IOException {
        List<String> out = new ArrayList<>();
        if (dinnerRows == null || dinnerRows.isEmpty()) return out;

        for (Row r : dinnerRows) {
            String label = readCellFlat(r, 0);
            if (label.contains("일자/요일") || label.contains("열량/단백질")) break;

            String cell = readCellKeepNewlines(r, col); // safe even if col >= size
            out.addAll(splitMenusByLineThenSlash(cell));
        }
        out = reattachLineSplitAllergens(out);
        out.removeIf(s -> s == null || s.isBlank());
        return out;
    }

    // ===== helpers =====
    private static boolean isDinnerLabel(String firstCol) {
        if (firstCol == null) return false;
        String t = firstCol.replaceAll("\\s+", ""); // spaces + newlines collapsed
        return t.contains("저녁") || t.contains("저녁간식") || t.contains("저녘") || t.contains("저녀");
    }

    private static List<String> reattachLineSplitAllergens(List<String> items) {
        List<String> out = new ArrayList<>();
        if (items == null) return out;
        for (String it : items) {
            if (it == null) continue;
            String t = it.trim();
            if (t.isEmpty()) continue;
            if (ALLERGENS.matcher(t).matches()) {
                if (!out.isEmpty()) {
                    String prev = out.get(out.size() - 1);
                    if (prev == null) prev = "";
                    out.set(out.size() - 1, prev + t);
                } else {
                    out.add(t);
                }
            } else {
                out.add(t);
            }
        }
        return out;
    }

    private static int effectiveMenuCol(int headerColIndex, List<String> header) {
        int usableOrdinal = 0;
        for (int i = 1; i <= headerColIndex && i < header.size(); i++) {
            String h = stripSpacesFlat(header.get(i));
            if (h == null || h.isEmpty()) continue;
            Matcher m = DAY_CELL.matcher(h);
            if (m.find()) usableOrdinal = usableOrdinal + 1;
        }
        int mapped = usableOrdinal; // body columns are 1-based after label col
        if (mapped < 1) mapped = headerColIndex; // fallback
        return mapped;
    }

    private static int findFirstUsableHeaderIndex(List<String> header) {
        if (header == null || header.size() <= 1) return -1;
        for (int i = 1; i < header.size(); i++) {
            String h = stripSpacesFlat(header.get(i));
            if (h == null || h.isEmpty()) continue;
            Matcher m = DAY_CELL.matcher(h);
            if (m.find()) return i;
        }
        return -1;
    }

    private static boolean isHolidayHeader(String headerFlat) {
        if (headerFlat == null) return false;
        String h = headerFlat.toLowerCase(Locale.ROOT).replace(" ", "");
        if (h.contains("대체휴일")) return true;
        if (h.contains("공휴일")) return true;
        if (h.contains("휴원")) return true;
        if (h.contains("휴무")) return true;
        if (h.contains("현충일")) return true;
        if (h.contains("어린이날")) return true;
        if (h.contains("석가탄신일")) return true;
        if (h.contains("광복절")) return true;
        if (h.contains("개천절")) return true;
        if (h.contains("한글날")) return true;
        if (h.contains("신정")) return true;
        if (h.contains("설날")) return true;
        if (h.contains("추석")) return true;
        if (h.contains("추석연휴")) return true;
        if (h.contains("성탄절")) return true;
        return false;
    }

    private static int[] parseYearMonthFromFilename(String fileName) {
        Pattern p = Pattern.compile("^(\\d{4})년\\s*(\\d{1,2})월\\s*시간연장형\\(만\\d+(?:-\\d+)?세\\)\\.hwp$");
        Matcher m = p.matcher(fileName);
        if (m.find()) {
            return new int[]{ Integer.parseInt(m.group(1)), Integer.parseInt(m.group(2)) };
        }
        return null;
    }

    private static String derivePrettyBasename(String fileName) {
        Pattern p = Pattern.compile("^(\\d{4})년\\s*(\\d{1,2})월\\s*시간연장형\\(만(\\d+(?:-\\d+)?)세\\)\\.hwp$");
        Matcher m = p.matcher(fileName);
        if (m.find()) {
            int yy = Integer.parseInt(m.group(1)) % 100;
            String yy2 = (yy < 10) ? "0" + yy : String.valueOf(yy);
            int mo = Integer.parseInt(m.group(2));
            String mm2 = (mo < 10) ? "0" + mo : String.valueOf(mo);
            return yy2 + "." + mm2 + ". 만" + m.group(3) + "세 시간연장형";
        }
        int dot = fileName.lastIndexOf('.');
        return (dot > 0) ? fileName.substring(0, dot) : fileName;
    }

    private static String safeDowKorean(int year, int month, int day) {
        try {
            DayOfWeek dow = LocalDate.of(year, month, day).getDayOfWeek();
            switch (dow) {
                case MONDAY:    return "월";
                case TUESDAY:   return "화";
                case WEDNESDAY: return "수";
                case THURSDAY:  return "목";
                case FRIDAY:    return "금";
                case SATURDAY:  return "토";
                case SUNDAY:    return "일";
            }
        } catch (Exception ignore) {}
        return null;
    }

    // ===== text extractors =====
    private static String cellText(Cell cell) throws IOException {
        ParagraphListInterface plis = cell.getParagraphList();
        if (plis == null || plis.getParagraphCount() == 0) return "";

        StringBuilder out = new StringBuilder();

        for (int i = 0; i < plis.getParagraphCount(); i++) {
            Paragraph p = plis.getParagraph(i);

            String line = fallbackPlainRuns(p);
            if (!line.isEmpty()) {
                if (out.length() > 0) out.append('\n');
                out.append(line);
            }

            if (p.getControlList() != null) {
                for (Control cc : p.getControlList()) {
                    if (cc.getType() == ControlType.Gso) {
                        GsoControl g = (GsoControl) cc;
                        Caption cap = g.getCaption();
                        if (cap != null && cap.getParagraphList() != null) {
                            ParagraphListInterface inner = cap.getParagraphList();
                            for (int j = 0; j < inner.getParagraphCount(); j++) {
                                String innerLine = fallbackPlainRuns(inner.getParagraph(j));
                                if (!innerLine.isEmpty()) {
                                    if (out.length() > 0) out.append('\n');
                                    out.append(innerLine);
                                }
                            }
                        }
                    }
                }
            }
        }
        return out.toString();
    }

    private static String fallbackPlainRuns(Paragraph p) throws UnsupportedEncodingException {
        if (p == null || p.getText() == null) return "";
        StringBuilder sb = new StringBuilder();
        for (HWPChar ch : p.getText().getCharList()) {
            if (ch.getType() == HWPCharType.Normal) sb.append(((HWPCharNormal) ch).getCh());
            else if (ch.getType() == HWPCharType.ControlInline) sb.append(' ');
        }
        return sb.toString();
    }

    private static String stripSpacesKeepNewlines(String s) {
        if (s == null) return "";
        String t = s
                .replace('\u00A0', ' ')
                .replace('\u2000', ' ').replace('\u2001', ' ').replace('\u2002', ' ')
                .replace('\u2003', ' ').replace('\u2004', ' ').replace('\u2005', ' ')
                .replace('\u2006', ' ').replace('\u2007', ' ').replace('\u2008', ' ')
                .replace('\u2009', ' ').replace('\u200A', ' ').replace('\u202F', ' ')
                .replace('\u205F', ' ').replace('\u3000', ' ');
        StringBuilder out = new StringBuilder();
        for (String line : t.split("\\r?\\n")) {
            String compact = SPACES_ONLY.matcher(line).replaceAll(" ").trim();
            if (compact.isEmpty()) continue;
            out.append(compact).append('\n');
        }
        if (out.length() == 0) return "";
        out.setLength(out.length() - 1);
        return out.toString();
    }

    private static String stripSpacesFlat(String s) {
        if (s == null) return "";
        String t = s
                .replace('\n', ' ')
                .replace('\r', ' ')
                .replace('\u00A0', ' ')
                .replace('\u2000', ' ').replace('\u2001', ' ').replace('\u2002', ' ')
                .replace('\u2003', ' ').replace('\u2004', ' ').replace('\u2005', ' ')
                .replace('\u2006', ' ').replace('\u2007', ' ').replace('\u2008', ' ')
                .replace('\u2009', ' ').replace('\u200A', ' ').replace('\u202F', ' ')
                .replace('\u205F', ' ').replace('\u3000', ' ');
        return SPACES_ONLY.matcher(t).replaceAll(" ").trim();
    }

    private static List<String> splitMenusByLineThenSlash(String rawWithNewlines) {
        List<String> out = new ArrayList<>();
        String cleaned = stripSpacesKeepNewlines(rawWithNewlines);
        if (cleaned.isEmpty()) return out;

        for (String line : cleaned.split("\\r?\\n")) {
            String lineTrim = line.trim();
            if (lineTrim.isEmpty()) continue;

            boolean dropLine = false;
            for (String hint : DROP_LINE_HINTS) {
                if (lineTrim.contains(hint)) { dropLine = true; break; }
            }
            if (dropLine) continue;

            for (String part : lineTrim.split("/")) {
                String v = part.trim();
                if (v.isEmpty()) continue;
                if (DROP_TOKENS_MISC.contains(v)) continue;
                if (DROP_TOKENS_FOODS.contains(v)) continue;
                out.add(v);
            }
        }
        return out;
    }

    private static String readCellFlat(Row row, int col) throws IOException {
        if (row == null || row.getCellList() == null) return "";
        List<Cell> cells = row.getCellList();
        if (col < 0 || col >= cells.size()) return "";
        return stripSpacesFlat(cellText(cells.get(col)));
    }

    private static String readCellKeepNewlines(Row row, int col) throws IOException {
        if (row == null || row.getCellList() == null) return "";
        List<Cell> cells = row.getCellList();
        if (col < 0 || col >= cells.size()) return "";
        return stripSpacesKeepNewlines(cellText(cells.get(col)));
    }

    private static List<String> collectRowTextsFlat(Row row) throws IOException {
        List<String> vals = new ArrayList<>();
        if (row.getCellList() == null) return vals;
        for (int i = 0; i < row.getCellList().size(); i++) vals.add(readCellFlat(row, i));
        return vals;
    }

    private static void dumpTableTsv(List<Row> rows, Path outTsv) {
        try (PrintWriter pw = new PrintWriter(Files.newBufferedWriter(outTsv))) {
            for (Row r : rows) {
                List<String> cols = new ArrayList<>();
                if (r.getCellList() != null) {
                    for (int i = 0; i < r.getCellList().size(); i++) {
                        String keep = readCellKeepNewlines(r, i).replace('\t', ' ');
                        cols.add(keep.replace("\n", "⏎"));
                    }
                }
                pw.println(String.join("\t", cols));
            }
        } catch (Exception ignore) {}
    }

    // New: Convert with explicit output path
    public static Path convert(Path inputHwpPath, Path outputJsonPath) throws Exception {
        if (inputHwpPath == null) throw new IllegalArgumentException("inputHwpPath is null");
        if (outputJsonPath == null) throw new IllegalArgumentException("outputJsonPath is null");

        // Ensure output dir exists
        Path parent = outputJsonPath.toAbsolutePath().getParent();
        if (parent != null) Files.createDirectories(parent);

        int[] ym = parseYearMonthFromFilename(inputHwpPath.getFileName().toString());
        if (ym != null) { PARSED_YEAR = ym[0]; PARSED_MONTH = ym[1]; }
        else { PARSED_YEAR = null; PARSED_MONTH = null; }

        String inAbs = inputHwpPath.toAbsolutePath().toString();
        HWP = HWPReader.fromFile(inAbs);

        Map<Integer, DayPlan> byDate = new LinkedHashMap<>();

        // scan all sections/paragraphs; process ControlType.Table only
        for (int s = 0; s < HWP.getBodyText().getSectionList().size(); s++) {
            Section sec = HWP.getBodyText().getSectionList().get(s);
            for (int p = 0; p < sec.getParagraphCount(); p++) {
                Paragraph para = sec.getParagraph(p);
                if (para.getControlList() == null) continue;
                for (Control c : para.getControlList()) {
                    if (c.getType() != ControlType.Table) continue;
                    ControlTable ct = (ControlTable) c;
                    // outTsv는 convert에서는 만들지 않고 null로 넘겨도 됨(현재 scanTable에서 실제로 사용 안 함)
                    scanTable(ct, byDate, null);
                }
            }
        }

        // finalize: ensure non-empty menus
        for (DayPlan dp : byDate.values()) {
            if (dp.menus == null || dp.menus.isEmpty()) dp.menus = List.of("없음");
        }

        List<DayPlan> list = new ArrayList<>(byDate.values());
        list.sort(Comparator.comparingInt(dp -> dp.date));

        new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT)
                .writeValue(outputJsonPath.toFile(), list);

        return outputJsonPath;
    }



    // English comment: Convert HWP to JSON string without writing intermediate JSON file
    public static String convertToJsonString(File hwpFile) throws Exception {
        if (hwpFile == null) throw new IllegalArgumentException("hwpFile is null");
        return convertToJsonString(hwpFile.toPath());
    }

    // English comment: Convert HWP to JSON string without writing intermediate JSON file
    public static String convertToJsonString(Path inputHwpPath) throws Exception {
        if (inputHwpPath == null) throw new IllegalArgumentException("inputHwpPath is null");

        int[] ym = parseYearMonthFromFilename(inputHwpPath.getFileName().toString());
        if (ym != null) { PARSED_YEAR = ym[0]; PARSED_MONTH = ym[1]; }
        else { PARSED_YEAR = null; PARSED_MONTH = null; }

        String inAbs = inputHwpPath.toAbsolutePath().toString();
        HWP = HWPReader.fromFile(inAbs);

        Map<Integer, DayPlan> byDate = new LinkedHashMap<>();

        // English comment: Keep the original table scanning logic (same as convert())
        for (int s = 0; s < HWP.getBodyText().getSectionList().size(); s++) {
            Section sec = HWP.getBodyText().getSectionList().get(s);
            for (int p = 0; p < sec.getParagraphCount(); p++) {
                Paragraph para = sec.getParagraph(p);
                if (para.getControlList() == null) continue;
                for (Control c : para.getControlList()) {
                    if (c.getType() != ControlType.Table) continue;
                    ControlTable ct = (ControlTable) c;
                    scanTable(ct, byDate, null);
                }
            }
        }

        List<DayPlan> list = new ArrayList<>(byDate.values());
        list.sort(Comparator.comparingInt(dp -> dp.date));

        return new ObjectMapper()
                .enable(SerializationFeature.INDENT_OUTPUT)
                .writeValueAsString(list);
    }

}