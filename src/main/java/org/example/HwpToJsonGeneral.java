package org.example;

import java.io.File;

import java.time.DayOfWeek;
import java.time.LocalDate;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

// ==== hwplib ====
import kr.dogfoot.hwplib.object.HWPFile;
import kr.dogfoot.hwplib.object.bodytext.Section;
import kr.dogfoot.hwplib.object.bodytext.ParagraphListInterface;
import kr.dogfoot.hwplib.object.bodytext.control.Control;
import kr.dogfoot.hwplib.object.bodytext.control.ControlType;
import kr.dogfoot.hwplib.object.bodytext.control.ControlTable;
import kr.dogfoot.hwplib.object.bodytext.control.gso.GsoControl;
import kr.dogfoot.hwplib.object.bodytext.control.gso.caption.Caption;
import kr.dogfoot.hwplib.object.bodytext.control.table.Cell;
import kr.dogfoot.hwplib.object.bodytext.control.table.Row;
import kr.dogfoot.hwplib.object.bodytext.paragraph.Paragraph;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.HWPChar;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.HWPCharNormal;
import kr.dogfoot.hwplib.object.bodytext.paragraph.text.HWPCharType;
import kr.dogfoot.hwplib.reader.HWPReader;
// ===============

import java.io.IOException;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class HwpToJsonGeneral {
    // ===== Default input path (used when no CLI arg is provided) =====
    private static final String DEFAULT_IN = "input/2026년 1월 일반형(만1-2세).hwp";

    private static final ObjectMapper MAPPER = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);
    private static HWPFile HWP;

    // Verbose logging
    private static final boolean DBG = true;

    private static Integer PARSED_YEAR  = null;
    private static Integer PARSED_MONTH = null;

    // ===== regex & tokens =====
    private static final Pattern DAY_CELL = Pattern.compile("^(\\d{1,2})\\((.)\\)");
    private static final Pattern SPACES_ONLY = Pattern.compile("[ \\t\\u00A0\\u2000-\\u200B\\u202F\\u205F\\u3000]+");

    private static final Set<String> DROP_TOKENS_MISC = Set.of("저 녁", "열량/단백질", "일자/요일", "없음", "오전간식", "점 심", "점심", "오후간식");
    private static final Set<String> DROP_TOKENS_FOODS = Set.of("", " ", "\\");

    private static final Pattern ALLERGENS      = Pattern.compile("^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲]+$");
    private static final Pattern ALLERGENS_TAIL = Pattern.compile("[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲]+$");
    private static final Pattern STEP_PREFIX    = Pattern.compile("(?i)^step\\s*[123]");

    private static final List<String> DROP_LINE_HINTS = List.of(
            "★ 푸드브릿지","푸드브릿지란","<식품 알레르기","원 산 지 표 시",
            "수산물","수산물가공품","식육가공품","가공품","국가명(",
            "[12월 제철 식재료]","[11월 제철 식재료]"
    );

    // ===== output model =====
    public static class DayPlan {
        public Integer date;      // 1..31; null for labeled entries (e.g., birthday)
        public String  dateStr;   // "birthday" when column is ★Birthday식단
        public String  weekday;   // optional; null for birthday
        public List<String> menus = new ArrayList<>();
        public List<String> pmDesert = new ArrayList<>();

        // Number of morning snack menus actually used for that day
        public Integer morningCount;

        public DayPlan() {}
        DayPlan(int d, String w) { date = d; weekday = w; }
        DayPlan(String label) { dateStr = label; }
    }

    public static void main(String[] args) throws Exception {
        String inPathStr;
        if (args != null && args.length > 0 && args[0] != null && !args[0].isBlank()) {
            inPathStr = args[0];
        } else {
            inPathStr = DEFAULT_IN;
        }
        Path inPath = Paths.get(inPathStr).toAbsolutePath();

        String base = derivePrettyBasename(inPath.getFileName().toString());

        int[] ym = parseYearMonthFromFilename(inPath.getFileName().toString());
        if (ym != null) {
            PARSED_YEAR  = ym[0];
            PARSED_MONTH = ym[1];
        }

        Path outDir = Paths.get("output").toAbsolutePath();
        Files.createDirectories(outDir);

        // If args[1] exists -> use it as OUT_JSON
        Path OUT_JSON;
        if (args != null && args.length >= 2 && args[1] != null && !args[1].isBlank()) {
            OUT_JSON = Paths.get(args[1].trim()).toAbsolutePath();
            if (OUT_JSON.getParent() != null) Files.createDirectories(OUT_JSON.getParent());
        } else {
            OUT_JSON = outDir.resolve(base + ".json");
        }

        // Debug TSV next to OUT_JSON
        String jsonName = OUT_JSON.getFileName().toString();
        String stem = jsonName;
        int dot = stem.lastIndexOf('.');
        if (dot > 0) stem = stem.substring(0, dot);
        Path OUT_TSV = OUT_JSON.getParent().resolve(stem + "-debug.tsv");

        System.out.println("[INPUT ] " + inPath);
        System.out.println("[OUTPUT] " + OUT_JSON);
        System.out.println("[DEBUG ] " + OUT_TSV);

        HWP = HWPReader.fromFile(inPath.toString());

        Map<Integer, DayPlan> byDate = new LinkedHashMap<>();
        DayPlan birthday = null;

        for (int s = 0; s < HWP.getBodyText().getSectionList().size(); s++) {
            Section sec = HWP.getBodyText().getSectionList().get(s);
            for (int p = 0; p < sec.getParagraphCount(); p++) {
                Paragraph para = sec.getParagraph(p);
                if (para.getControlList() == null) continue;
                for (Control c : para.getControlList()) {
                    if (c.getType() != ControlType.Table) continue;
                    ControlTable ct = (ControlTable) c;
                    DayPlan found = scanTable(ct, byDate, OUT_TSV);
                    if (found != null) birthday = found;
                }
            }
        }

        dumpByDateSnapshot(byDate);

        List<DayPlan> list = new ArrayList<>(byDate.values());
        list.sort(Comparator.comparingInt(dp -> dp.date == null ? Integer.MAX_VALUE : dp.date));
        for (DayPlan dp : list) {
            if (dp.menus == null || dp.menus.isEmpty()) dp.menus = List.of("없음");
            if (dp.pmDesert == null || dp.pmDesert.isEmpty()) dp.pmDesert = List.of("없음");
        }
        if (birthday != null) {
            if (birthday.menus == null || birthday.menus.isEmpty()) birthday.menus = List.of("없음");
            if (birthday.pmDesert == null || birthday.pmDesert.isEmpty()) birthday.pmDesert = List.of("없음");
        }

        List<DayPlan> finalOut = new ArrayList<>(list);
        if (birthday != null) finalOut.add(birthday);

        System.out.println("==== FINAL JSON PREVIEW ====");
        for (DayPlan dp : finalOut) {
            if (dp.dateStr != null) {
                System.out.println(dp.dateStr + ": MENUS=" + dp.menus + " | PM=" + dp.pmDesert + " | morningCount=" + dp.morningCount);
            } else {
                System.out.println(dp.date + "(" + dp.weekday + "): MENUS=" + dp.menus + " | PM=" + dp.pmDesert + " | morningCount=" + dp.morningCount);
            }
        }

        MAPPER.writeValue(OUT_JSON.toFile(), finalOut);
        System.out.println("[WRITE] " + OUT_JSON);

        List<DayPlan> readBack = MAPPER.readValue(OUT_JSON.toFile(), new TypeReference<List<DayPlan>>() {});
        System.out.println("==== READ-BACK PREVIEW ====");
        for (DayPlan dp : readBack) {
            if (dp.dateStr != null) {
                System.out.println(dp.dateStr + ": MENUS=" + dp.menus + " | PM=" + dp.pmDesert + " | morningCount=" + dp.morningCount);
            } else {
                System.out.println(dp.date + "(" + dp.weekday + "): MENUS=" + dp.menus + " | PM=" + dp.pmDesert + " | morningCount=" + dp.morningCount);
            }
        }
        String inMem = MAPPER.writeValueAsString(finalOut);
        String onDisk = MAPPER.writeValueAsString(readBack);
        if (!inMem.equals(onDisk)) {
            System.out.println("[WARN] In-memory and on-disk JSON differ!");
        } else {
            System.out.println("[OK] File content matches preview.");
        }
        auditSaturdays(byDate);
    }

    // Derive pretty basename from filename
    private static String derivePrettyBasename(String fileName) {
        Pattern p = Pattern.compile("^(\\d{4})년\\s*(\\d{1,2})월\\s*일반형\\(만(\\d+(?:-\\d+)?)세\\)\\.hwp$");
        Matcher m = p.matcher(fileName);
        if (m.find()) {
            int yy = Integer.parseInt(m.group(1)) % 100;
            String yy2 = (yy < 10) ? "0" + yy : String.valueOf(yy);
            int month = Integer.parseInt(m.group(2));
            String mm2 = (month < 10) ? "0" + month : String.valueOf(month);
            return yy2 + "." + mm2 + ". " + "만" + m.group(3) + "세 일반형";
        }
        int dot = fileName.lastIndexOf('.');
        if (dot > 0) return fileName.substring(0, dot);
        return fileName;
    }

    // ====== Scan a table control ======
    private static DayPlan scanTable(ControlTable ct, Map<Integer, DayPlan> byDate, Path OUT_TSV) throws IOException {
        List<Row> rows = ct.getRowList();
        if (rows == null || rows.isEmpty()) return null;

//            dumpTableTsv(rows, OUT_TSV); // debug snapshot per table

        List<String> currentHeader = null;
        List<Row>    menuRows      = new ArrayList<>();
        DayPlan      birthday      = null;

        for (int r = 0; r < rows.size(); r++) {
            String first = readCellFlat(rows.get(r), 0);
            boolean isHeader = first.contains("일자/요일");

            if (isHeader) {
                if (currentHeader != null && blockHasMorning(menuRows) && hasUsableHeader(currentHeader)) {
                    DayPlan b = flushBlock(currentHeader, menuRows, byDate);
                    if (b != null) birthday = b;
                    menuRows.clear();
                } else {
                    if (currentHeader != null) {
                        if (!blockHasMorning(menuRows)) System.out.println("[SKIP-FLUSH] previous block had no '오전간식' row.");
                        else if (!hasUsableHeader(currentHeader)) System.out.println("[SKIP-FLUSH] previous block header had NO usable date/birthday column.");
                    }
                    menuRows.clear();
                }
                currentHeader = collectRowTextsFlat(rows.get(r));
                System.out.println("=== HEADER FOUND @ row " + r + " ===");
                for (int i = 0; i < currentHeader.size(); i++) {
                    System.out.println("  header[" + i + "] = [" + currentHeader.get(i) + "]");
                }
                continue;
            }

            if (currentHeader != null) {
                if (first.contains("열량/단백질") || first.contains("일자/요일")) {
                    if (blockHasMorning(menuRows) && hasUsableHeader(currentHeader)) {
                        DayPlan b = flushBlock(currentHeader, menuRows, byDate);
                        if (b != null) birthday = b;
                    } else {
                        System.out.println("[SKIP-FLUSH-INNER] blockHasMorning=" + blockHasMorning(menuRows)
                                + ", hasUsableHeader=" + hasUsableHeader(currentHeader));
                    }
                    menuRows.clear();
                    continue;
                }
                menuRows.add(rows.get(r));
            }
        }

        if (currentHeader != null && blockHasMorning(menuRows) && hasUsableHeader(currentHeader)) {
            DayPlan b = flushBlock(currentHeader, menuRows, byDate);
            if (b != null) birthday = b;
        } else if (currentHeader != null) {
            System.out.println("[SKIP-FLUSH-LAST] blockHasMorning=" + blockHasMorning(menuRows)
                    + ", hasUsableHeader=" + hasUsableHeader(currentHeader));
        }
        return birthday;
    }

    private static boolean hasUsableHeader(List<String> header) {
        int idx = findFirstUsableHeaderIndex(header);
        System.out.println("[HEADER-USABLE] first usable index = " + idx);
        return idx >= 1;
    }

    private static boolean blockHasMorning(List<Row> menuRows) throws IOException {
        if (menuRows == null) return false;
        for (int i = 0; i < menuRows.size(); i++) {
            String first = readCellFlat(menuRows.get(i), 0);
            if (first.contains("오전간식")) return true;
        }
        return false;
    }

    /**
     * Branch D (holiday-streak queue) with morning-only queue.
     * - On holiday cells: enqueue morning-only chunk and non-morning lunch/dinner chunk separately.
     * - On consumption cells (normal days, and Saturday also consumes):
     *   * Morning: use today's morningPart if present; otherwise consume one from morning-queue.
     *   * MENUS rest: consume exactly one queued rest-chunk (FIFO) if present; push today's rest as a new chunk.
     *   * PM: consume one queued PM chunk if present; otherwise use today's PM; then enqueue today's PM.
     */
    private static DayPlan flushBlock(
            List<String> currentHeader,
            List<Row> menuRows,
            Map<Integer, DayPlan> byDate
    ) throws IOException {
        if (currentHeader == null || currentHeader.isEmpty()) return null;
        if (menuRows == null) return null;

        System.out.println("\n--- FLUSH BLOCK ---");
        System.out.println("header size = " + currentHeader.size());
        for (int i = 0; i < currentHeader.size(); i++) {
            System.out.println("  H[" + i + "] = [" + currentHeader.get(i) + "]");
        }
        System.out.println("menuRows = " + menuRows.size());

        int morningIdx = findRowByLabel(menuRows, "오전간식");
        int afternoonIdx = findRowByLabel(menuRows, "오후간식");

        System.out.println("morningIdx = " + morningIdx);
        if (morningIdx >= 0 && morningIdx < menuRows.size()) {
            System.out.println("morning label cell = [" + readCellFlat(menuRows.get(morningIdx), 0) + "]");
        }

        int cStart = findFirstUsableHeaderIndex(currentHeader);
        if (cStart < 1) {
            System.out.println("No usable header column found. Skip this block.");
            return null;
        }
        System.out.println("usable header starts at index = " + cStart + " (H[" + cStart + "])");

        int colCount = currentHeader.size();

        // --- holiday-streak queues (day-chunks) ---
        Deque<List<String>> carryMenusQ   = new ArrayDeque<>(); // FIFO of non-morning lunch/dinner chunks
        Deque<List<String>> carryPmQ      = new ArrayDeque<>(); // FIFO of PM chunks
        Deque<List<String>> carryMorningQ = new ArrayDeque<>(); // FIFO of morning-only chunks

        DayPlan birthday = null;

        for (int c = cStart; c < colCount; c++) {
            String headerTextRaw = currentHeader.get(c);
            String headerText = stripSpacesFlat(headerTextRaw);

            int effC = effectiveMenuCol(c, currentHeader);

            boolean isBirthdayCol = headerText.toLowerCase(Locale.ROOT).contains("birthday");
            if (isBirthdayCol) {
                System.out.println("[COL] H[" + c + "] BIRTHDAY-COL = [" + headerTextRaw + "]");
                List<String> merged = collectMergedForColumnSkippingPM(menuRows, effC, morningIdx, afternoonIdx);
                merged = reattachLineSplitAllergens(merged);
                merged = mergeTrailingAllergenTokens(merged);

                if (birthday == null) birthday = new DayPlan("birthday");
                if (merged != null && !merged.isEmpty()) {
                    if (birthday.menus == null || birthday.menus.isEmpty() || birthday.menus.equals(List.of("없음"))) {
                        birthday.menus = new ArrayList<>();
                    }
                    birthday.menus.addAll(merged);
                }
                if (afternoonIdx >= 0 && afternoonIdx < menuRows.size()) {
                    String pmCell = readCellKeepNewlines(menuRows.get(afternoonIdx), effC);
                    List<String> pmParts = splitMenusByLineThenSlash(pmCell);
                    pmParts = reattachLineSplitAllergens(pmParts);
                    pmParts = mergeTrailingAllergenTokens(pmParts);
                    if (birthday.pmDesert == null || birthday.pmDesert.isEmpty() || birthday.pmDesert.equals(List.of("없음"))) {
                        birthday.pmDesert = new ArrayList<>();
                    }
                    if (pmParts != null && !pmParts.isEmpty()) birthday.pmDesert.addAll(pmParts);
                }
                continue;
            }

            Matcher m = DAY_CELL.matcher(headerText);
            if (!m.find()) {
                System.out.println("[COL] H[" + c + "] NOT A DATE HEADER -> [" + headerTextRaw + "]");
                continue;
            }
            int    date          = Integer.parseInt(m.group(1));
            String headerWeekday = m.group(2);

            String weekday;
            String computedDow = null;
            if (PARSED_YEAR != null && PARSED_MONTH != null) {
                String calc = safeDowKorean(PARSED_YEAR, PARSED_MONTH, date);
                computedDow = calc;
                if (calc != null) {
                    if ("일".equals(calc)) {
                        System.out.println("[SKIP] Sunday skip at " + date + " (computed DoW=일).");
                        continue;
                    }
                    if (!calc.equals(headerWeekday)) {
                        System.out.println("[SKIP] DoW mismatch at " + date + ": header=" + headerWeekday + ", computed=" + calc);
                        continue;
                    }
                    weekday = calc;
                } else {
                    weekday = headerWeekday;
                }
            } else {
                weekday = headerWeekday;
            }

            DayPlan dp = byDate.computeIfAbsent(date, d -> new DayPlan(d, weekday));

            System.out.println("\n====== COL START ======");
            System.out.println("COL " + date + "(" + weekday + "), headerRaw=[" + headerTextRaw + "] computedDoW=" + computedDow);

            boolean allBlank = isTrulyEmptyColumn(menuRows, effC);
            boolean isHoliday = isHolidayHeader(headerText);
            boolean morningSegmentBlank = isMorningSegmentBlank(menuRows, effC, morningIdx);

            boolean applyHolidayRule = isHoliday;

            System.out.println("allBlankColumn=" + allBlank);
            System.out.println("isHoliday=" + isHoliday + ", applyHolidayRule=" + applyHolidayRule
                    + ", morningSegmentBlank=" + morningSegmentBlank);

            List<String> mergedAll   = new ArrayList<>();
            List<String> morningPart = new ArrayList<>();
            List<String> restParts   = new ArrayList<>();
            List<String> pmParts     = new ArrayList<>();

            boolean morningSeen     = false;
            int blankAfterMorning   = 0;

            for (int ri = 0; ri < menuRows.size(); ri++) {
                Row mr       = menuRows.get(ri);
                String label = readCellFlat(mr, 0);
                String cell  = readCellKeepNewlines(mr, effC);

                if (label.contains("오전간식")) {
                    if (morningSeen) break;
                    morningSeen = true;
                    blankAfterMorning = 0;
                }

                if (morningSeen) {
                    String labelTrimmed = label
                            .replaceAll("[\\s\\u00A0\\u2000-\\u200B\\u202F\\u205F\\u3000]", "")
                            .replaceAll("\\r?\\n", "")
                            .trim();
                    if (labelTrimmed.isEmpty()) {
                        blankAfterMorning = blankAfterMorning + 1;
                        if (blankAfterMorning >= 2) break;
                        continue;
                    } else {
                        blankAfterMorning = 0;
                    }
                }

                if (label.contains("없음") || label.contains("성탄절") || label.contains("Family")) break;

                if (label.contains("오후간식")) {
                    List<String> partsPM = splitMenusByLineThenSlash(cell);
                    if (!partsPM.isEmpty()) pmParts.addAll(partsPM);
                    continue;
                }

                List<String> parts = splitMenusByLineThenSlash(cell);
                if (ri == morningIdx) {
                    if (!parts.isEmpty()) {
                        morningPart.addAll(parts);
                        mergedAll.addAll(parts);
                    }
                } else {
                    if (!parts.isEmpty()) {
                        restParts.addAll(parts);
                        mergedAll.addAll(parts);
                    }
                }
            }

            mergedAll.removeIf(x -> x == null || x.isBlank());
            morningPart.removeIf(x -> x == null || x.isBlank());
            restParts.removeIf(x -> x == null || x.isBlank());
            pmParts.removeIf(x -> x == null || x.isBlank());

            mergedAll = reattachLineSplitAllergens(mergedAll);
            pmParts   = reattachLineSplitAllergens(pmParts);

            System.out.println("mergedAll.size=" + mergedAll.size() + " mergedAll=" + mergedAll);
            System.out.println("pmParts=" + pmParts);

            // ====== holiday-streak enqueue only ======
            if (applyHolidayRule) {
                if (!morningPart.isEmpty()) carryMorningQ.addLast(new ArrayList<>(morningPart));
                if (!restParts.isEmpty())   carryMenusQ.addLast(new ArrayList<>(restParts));
                if (!pmParts.isEmpty())     carryPmQ.addLast(new ArrayList<>(pmParts));

                dp.menus = List.of("없음");
                dp.pmDesert = List.of("없음");
                dp.morningCount = Integer.valueOf(0);

                System.out.println("HOLIDAY ENQUEUE -> morningQ=" + carryMorningQ.size()
                        + ", menusQ=" + carryMenusQ.size() + ", pmQ=" + carryPmQ.size());
                System.out.println("====== COL END ======");
                continue;
            }

            // ====== CONSUMPTION DAY ======
            List<String> morningUse = new ArrayList<>();
            if (!morningPart.isEmpty()) {
                morningUse.addAll(morningPart);
            } else if (!carryMorningQ.isEmpty()) {
                List<String> q = carryMorningQ.pollFirst();
                if (q != null && !q.isEmpty()) morningUse.addAll(q);
            }

            List<String> finalMenus = new ArrayList<>();
            if (!morningUse.isEmpty()) finalMenus.addAll(morningUse);

            if (!carryMenusQ.isEmpty()) {
                List<String> chunk = carryMenusQ.pollFirst();
                if (chunk != null && !chunk.isEmpty()) finalMenus.addAll(chunk);
                if (!restParts.isEmpty()) carryMenusQ.addLast(new ArrayList<>(restParts));
            } else {
                if (!restParts.isEmpty()) finalMenus.addAll(restParts);
            }

            finalMenus = mergeTrailingAllergenTokens(finalMenus);

            if (dp.menus.equals(List.of("없음"))) dp.menus = new ArrayList<>();
            if (!finalMenus.isEmpty()) dp.menus.addAll(finalMenus);
            else dp.menus = List.of("없음");

            // Set morningCount based on number of morning menus used
            if (morningUse.isEmpty()) dp.morningCount = Integer.valueOf(0);
            else dp.morningCount = Integer.valueOf(morningUse.size());

            // PM: consume one chunk if available; enqueue today's pm
            List<String> pmUse;
            if (!carryPmQ.isEmpty()) {
                pmUse = carryPmQ.pollFirst();
                if (!pmParts.isEmpty()) carryPmQ.addLast(new ArrayList<>(pmParts));
            } else {
                pmUse = pmParts;
            }
            if (pmUse == null || pmUse.isEmpty()) {
                dp.pmDesert = List.of("없음");
            } else {
                dp.pmDesert = mergeTrailingAllergenTokens(pmUse);
            }

            System.out.println("CONSUME -> dp[" + date + "].menus=" + dp.menus + " pm=" + dp.pmDesert
                    + " | morningCount=" + dp.morningCount
                    + " | queues: morningQ=" + carryMorningQ.size()
                    + ", menusQ=" + carryMenusQ.size() + ", pmQ=" + carryPmQ.size());
            System.out.println("====== COL END ======");
        }
        return birthday;
    }

    // Map header column to real body column
    private static int effectiveMenuCol(int headerColIndex, List<String> header) {
        int usableOrdinal = 0;
        int i = 1;
        while (i <= headerColIndex && i < header.size()) {
            String h = stripSpacesFlat(header.get(i));
            boolean isUsable = false;
            if (h != null && !h.isEmpty()) {
                String low = h.toLowerCase(Locale.ROOT);
                if (low.contains("birthday")) isUsable = true;
                else {
                    Matcher m = DAY_CELL.matcher(h);
                    if (m.find()) isUsable = true;
                }
            }
            if (isUsable) usableOrdinal = usableOrdinal + 1;
            i = i + 1;
        }
        int mapped = usableOrdinal;
        if (mapped < 1) mapped = headerColIndex;
        return mapped;
    }

    private static int findFirstUsableHeaderIndex(List<String> header) {
        if (header == null || header.size() <= 1) return -1;
        int size = header.size();
        for (int i = 1; i < size; i++) {
            String h = stripSpacesFlat(header.get(i));
            if (h == null || h.isEmpty()) continue;
            String low = h.toLowerCase(Locale.ROOT);
            if (low.contains("birthday")) {
                System.out.println("[HEADER PICK] H[" + i + "] chosen as first usable (birthday): [" + header.get(i) + "]");
                return i;
            }
            Matcher m = DAY_CELL.matcher(h);
            if (m.find()) {
                System.out.println("[HEADER PICK] H[" + i + "] chosen as first usable (date): [" + header.get(i) + "]");
                return i;
            }
        }
        return -1;
    }

    private static List<String> collectMergedForColumnSkippingPM(List<Row> menuRows, int effC, int morningIdx, int afternoonIdx) throws IOException {
        List<String> mergedAll = new ArrayList<>();
        boolean morningSeen   = false;
        int blankAfterMorning = 0;

        for (int ri = 0; ri < menuRows.size(); ri++) {
            Row mr       = menuRows.get(ri);
            String label = readCellFlat(mr, 0);
            String cell  = readCellKeepNewlines(mr, effC);

            if (label.contains("오전간식")) {
                if (morningSeen) break;
                morningSeen = true;
                blankAfterMorning = 0;
            }

            if (morningSeen) {
                String labelTrimmed = label
                        .replaceAll("[\\s\\u00A0\\u2000-\\u200B\\u202F\\u205F\\u3000]", "")
                        .replaceAll("\\r?\\n", "")
                        .trim();
                if (labelTrimmed.isEmpty()) {
                    blankAfterMorning = blankAfterMorning + 1;
                    if (blankAfterMorning >= 2) break;
                    continue;
                } else {
                    blankAfterMorning = 0;
                }
            }

            if (label.contains("없음") || label.contains("성탄절") || label.contains("Family")) break;
            if (label.contains("오후간식")) continue;

            List<String> parts = splitMenusByLineThenSlash(cell);
            if (!parts.isEmpty()) mergedAll.addAll(parts);
        }
        mergedAll.removeIf(x -> x == null || x.isBlank());
        return mergedAll;
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

    // ===== whitespace helpers =====
    private static String stripSpacesKeepNewlines(String s) {
        if (s == null) return "";
        String t = s.replace('\u00A0',' ')
                .replace('\u2000',' ').replace('\u2001',' ').replace('\u2002',' ').replace('\u2003',' ')
                .replace('\u2004',' ').replace('\u2005',' ').replace('\u2006',' ').replace('\u2007',' ')
                .replace('\u2008',' ').replace('\u2009',' ').replace('\u200A',' ')
                .replace('\u202F',' ').replace('\u205F',' ').replace('\u3000',' ');
        StringBuilder out = new StringBuilder();
        for (String line : t.split("\\r?\\n")) {
            String compact = SPACES_ONLY.matcher(line).replaceAll(" ").trim();
            if (!compact.isEmpty()) out.append(compact).append('\n');
        }
        if (out.length() == 0) return "";
        out.setLength(out.length() - 1);
        return out.toString();
    }

    private static String stripSpacesFlat(String s) {
        if (s == null) return "";
        String t = s.replace('\n',' ').replace('\r',' ')
                .replace('\u00A0',' ')
                .replace('\u2000',' ').replace('\u2001',' ').replace('\u2002',' ').replace('\u2003',' ')
                .replace('\u2004',' ').replace('\u2005',' ').replace('\u2006',' ').replace('\u2007',' ')
                .replace('\u2008',' ').replace('\u2009',' ').replace('\u200A',' ')
                .replace('\u202F',' ').replace('\u205F',' ').replace('\u3000',' ');
        return SPACES_ONLY.matcher(t).replaceAll(" ").trim();
    }

    // ===== menu parsing =====
    private static List<String> splitMenusByLineThenSlash(String rawWithNewlines) {
        List<String> out = new ArrayList<>();
        String cleaned = stripSpacesKeepNewlines(rawWithNewlines);
        if (cleaned.isEmpty()) return out;

        for (String line : cleaned.split("\\r?\\n")) {
            String z = line.trim();
            if (z.isEmpty()) continue;

            boolean drop = false;
            for (String hint : DROP_LINE_HINTS) {
                if (z.contains(hint)) { drop = true; break; }
            }
            if (drop) continue;

            for (String part : z.split("/")) {
                String v = part.trim();
                if (v.isEmpty()) continue;
                if (DROP_TOKENS_MISC.contains(v)) continue;
                if (DROP_TOKENS_FOODS.contains(v)) continue;

                if (STEP_PREFIX.matcher(v).find()) {
                    v = STEP_PREFIX.matcher(v).replaceFirst("");
                    v = v.trim();
                }

                // If menu has pattern like "시리얼(모닝빵)", split into "시리얼" and "(모닝빵)"
                int idxParen = v.indexOf('(');
                if (idxParen > 0 && v.endsWith(")")) {
                    String left = v.substring(0, idxParen).trim();
                    String right = v.substring(idxParen).trim();
                    if (!left.isEmpty()) out.add(left);
                    if (!right.isEmpty()) out.add(right);
                    continue;
                }

                // If item starts with '&', attach to previous menu item
                if (v.startsWith("&") && !out.isEmpty()) {
                    int last = out.size() - 1;
                    out.set(last, out.get(last) + v);
                } else {
                    out.add(v);
                }
            }
        }
        return out;
    }

    private static List<String> reattachLineSplitAllergens(List<String> items) {
        List<String> out = new ArrayList<>();
        if (items == null || items.isEmpty()) return out;

        for (String raw : items) {
            if (raw == null) continue;
            String cur = raw.trim();
            if (cur.isEmpty()) continue;

            if (ALLERGENS.matcher(cur).matches()) {
                if (!out.isEmpty()) {
                    int last = out.size() - 1;
                    out.set(last, out.get(last) + cur);
                } else {
                    out.add(cur);
                }
                continue;
            }

            out.add(cur);
        }
        return out;
    }

    // ===== helpers =====
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

    private static int findRowByLabel(List<Row> rows, String keyword) throws IOException {
        if (rows == null) return -1;
        for (int i = 0; i < rows.size(); i++) {
            String first = readCellFlat(rows.get(i), 0);
            if (first.contains(keyword)) return i;
        }
        return -1;
    }

    private static boolean isTrulyEmptyColumn(List<Row> menuRows, int col) throws IOException {
        if (menuRows == null || menuRows.isEmpty()) return true;
        for (Row r : menuRows) {
            if (!isBlankMenuCell(r, col)) return false;
        }
        return true;
    }

    private static boolean isBlankMenuCell(Row row, int col) throws IOException {
        if (row == null || row.getCellList() == null) return true;
        List<Cell> cells = row.getCellList();
        if (col < 0 || col >= cells.size()) return true;

        String raw     = cellText(cells.get(col));
        String cleaned = stripSpacesKeepNewlines(raw);

        System.out.println("isBlankMenuCell: raw=[" + raw + "]");
        System.out.println("isBlankMenuCell: cleaned=[" + cleaned + "]");
        String cleaned2 = cleaned.replace("\n", " ").trim();
        System.out.println("isBlankMenuCell: cleaned2=[" + cleaned2 + "], len=" + cleaned2.length());

        return cleaned2.isEmpty();
    }

    private static boolean isMorningSegmentBlank(List<Row> menuRows, int col, int morningIdx) throws IOException {
        if (menuRows == null || menuRows.isEmpty()) return true;
        int start = morningIdx >= 0 ? morningIdx : 0;

        boolean sawAny = false;
        for (int i = start; i < menuRows.size(); i++) {
            String label = readCellFlat(menuRows.get(i), 0);

            String labelTrimmed = label
                    .replaceAll("[\\s\\u00A0\\u2000-\\u200B\\u202F\\u205F\\u3000]", "")
                    .replaceAll("\\r?\\n", "")
                    .trim();

            if (i > start && !labelTrimmed.isEmpty() && !label.contains("오전간식")) break;

            sawAny = true;
            if (!isBlankMenuCell(menuRows.get(i), col)) return false;
        }
        return sawAny;
    }

    // Stitch trailing allergen-only tokens to the previous item
    private static List<String> mergeTrailingAllergenTokens(List<String> in) {
        List<String> out = new ArrayList<>();
        if (in == null) return out;

        for (String raw : in) {
            if (raw == null) continue;
            String cur = raw.trim();
            if (cur.isEmpty()) continue;

            if (ALLERGENS.matcher(cur).matches()) {
                if (!out.isEmpty()) {
                    int last = out.size() - 1;
                    out.set(last, out.get(last) + cur);
                } else {
                    out.add(cur);
                }
                continue;
            }
            out.add(cur);
        }
        return out;
    }

    private static boolean isHolidayHeader(String headerFlat) {
        if (headerFlat == null) return false;
        String h = headerFlat.toLowerCase(Locale.ROOT).replace(" ", "");
        if (h.contains("공휴일")) return true;
        if (h.contains("Day")) return false;

        if (h.contains("공휴일")) return true;
        if (h.contains("휴원")) return true;
        if (h.contains("휴무")) return true;
        if (h.contains("선거")) return true;
        if (h.contains("현충일")) return true;
        if (h.contains("어린이날")) return true;
        if (h.contains("석가탄신일")) return true;
        if (h.contains("대체휴일")) return true;
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
        Pattern p = Pattern.compile("^(\\d{4})년\\s*(\\d{1,2})월\\s*일반형\\(만\\d+(?:-\\d+)?세\\)\\.hwp$");
        Matcher m = p.matcher(fileName);
        if (m.find()) {
            int y = Integer.parseInt(m.group(1));
            int mo = Integer.parseInt(m.group(2));
            return new int[]{ y, mo };
        }
        return null;
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

    private static List<String> collectRowTextsFlat(Row row) throws IOException {
        List<String> vals = new ArrayList<>();
        if (row.getCellList() == null) return vals;
        for (int i = 0; i < row.getCellList().size(); i++) {
            vals.add(readCellFlat(row, i));
        }
        return vals;
    }

//        private static void dumpTableTsv(List<Row> rows, Path OUT_TSV) {
//            try (PrintWriter pw = new PrintWriter(Files.newBufferedWriter(OUT_TSV))) {
//                for (Row r : rows) {
//                    List<String> cols = new ArrayList<>();
//                    if (r.getCellList() != null) {
//                        for (int i = 0; i < r.getCellList().size(); i++) {
//                            String keep = readCellKeepNewlines(r, i).replace('\t', ' ');
//                            cols.add(keep.replace("\n", "⏎"));
//                        }
//                    }
//                    pw.println(String.join("\t", cols));
//                }
//            } catch (Exception ignore) {}
//        }

    private static void dumpByDateSnapshot(Map<Integer, DayPlan> byDate) {
        System.out.println("\n==== SNAPSHOT byDate ====");
        for (Map.Entry<Integer, DayPlan> e : byDate.entrySet()) {
            Integer d = e.getKey();
            DayPlan dp = e.getValue();
            int msize = 0;
            if (dp != null && dp.menus != null) msize = dp.menus.size();
            System.out.println(" - " + d + "(" + dp.weekday + ") menus=" + msize + " morningCount=" + dp.morningCount);
        }
        System.out.println("==== SNAPSHOT END ====\n");
    }

    private static void auditSaturdays(Map<Integer, DayPlan> byDate) {
        System.out.println("\n==== SATURDAY AUDIT ====");
        if (PARSED_YEAR == null || PARSED_MONTH == null) {
            System.out.println("[SAT-AUDIT] year/month unknown -> skip audit.");
            return;
        }
        List<Integer> missing = new ArrayList<>();
        LocalDate base = LocalDate.of(PARSED_YEAR, PARSED_MONTH, 1);
        int len = base.lengthOfMonth();
        for (int day = 1; day <= len; day++) {
            LocalDate cur = LocalDate.of(PARSED_YEAR, PARSED_MONTH, day);
            if (cur.getDayOfWeek() == DayOfWeek.SATURDAY) {
                DayPlan dp = byDate.get(day);
                if (dp == null) {
                    System.out.println("!!! MISSING SATURDAY: " + day + " (토) -> not present in byDate");
                    missing.add(day);
                } else {
                    int msize = 0;
                    String sample = "";
                    if (dp.menus != null) {
                        msize = dp.menus.size();
                        sample = joinFirst(dp.menus, 3);
                    }
                    System.out.println("OK SAT: " + day + " (토) menus=" + msize + " sample=" + sample + " morningCount=" + dp.morningCount);
                }
            }
        }
        if (!missing.isEmpty()) {
            System.out.println("\n[SAT-AUDIT SUMMARY] Missing Saturdays: " + missing);
        } else {
            System.out.println("\n[SAT-AUDIT SUMMARY] All Saturdays present.");
        }
        System.out.println("==== SATURDAY AUDIT END ====\n");
    }

    private static boolean isBirthdayHeader(String s) {
        if (s == null) return false;
        String t = s.trim();
        if (t.contains("Birthday")) return true;
        if (t.contains("★Birthday")) return true;
        if (t.contains("★Birthday식단")) return true;
        return false;
    }

    private static boolean isDateHeader(String s) {
        if (s == null) return false;
        String t = s.trim();
        if (t.length() < 4) return false;
        int i = 0;
        while (i < t.length() && Character.isDigit(t.charAt(i))) i = i + 1;
        if (i == 0) return false;
        if (i >= t.length()) return false;
        if (t.charAt(i) != '(') return false;
        int j = i + 1;
        if (j >= t.length()) return false;
        char dow = t.charAt(j);
        if (dow != '월' && dow != '화' && dow != '수' && dow != '목' && dow != '금' && dow != '토' && dow != '일') return false;
        int k = j + 1;
        if (k >= t.length()) return false;
        if (t.charAt(k) != ')') return false;
        return true;
    }

    private static int pickFirstUsableHeader(java.util.List<String> headerCells, java.util.function.BiConsumer<String, String> debugLogKV, java.util.function.Consumer<String> debugLog) {
        int idx = -1;
        int n = headerCells.size();
        int i = 0;

        while (i < n) {
            String cell = headerCells.get(i);
            boolean isBD = isBirthdayHeader(cell);
            boolean isDT = isDateHeader(cell);
            if (isBD || isDT) {
                idx = i;
                break;
            }
            i = i + 1;
        }

        if (idx >= 0) {
            if (isBirthdayHeader(headerCells.get(idx))) debugLogKV.accept("[HEADER PICK]", "H[" + idx + "] chosen as first usable (birthday): [" + headerCells.get(idx) + "]");
            else debugLogKV.accept("[HEADER PICK]", "H[" + idx + "] chosen as first usable (date): [" + headerCells.get(idx) + "]");
            debugLogKV.accept("[HEADER-USABLE]", "first usable index = " + idx);
        } else {
            debugLog.accept("[HEADER PICK] No usable header found in this block (preface-only).");
        }
        return idx;
    }

    private static String joinFirst(List<String> items, int n) {
        if (items == null) return "";
        List<String> cut = new ArrayList<>();
        int limit = n;
        for (int i = 0; i < items.size(); i++) {
            if (i >= limit) break;
            String v = items.get(i);
            if (v != null) cut.add(v);
        }
        return String.join(", ", cut);
    }

    private static void dbg(String s) {
        if (DBG) System.out.println(s);
    }

    // New: Convert with explicit output path (for AllInOne)
    public static Path convert(Path inputHwpPath, Path outputJsonPath) throws Exception {
        if (inputHwpPath == null) throw new IllegalArgumentException("inputHwpPath is null");
        if (outputJsonPath == null) throw new IllegalArgumentException("outputJsonPath is null");

        Path inPath = inputHwpPath.toAbsolutePath();

        int[] ym = parseYearMonthFromFilename(inPath.getFileName().toString());
        if (ym != null) {
            PARSED_YEAR = ym[0];
            PARSED_MONTH = ym[1];
        } else {
            PARSED_YEAR = null;
            PARSED_MONTH = null;
        }

        Path parent = outputJsonPath.toAbsolutePath().getParent();
        if (parent != null) Files.createDirectories(parent);

        HWP = HWPReader.fromFile(inPath.toString());

        Map<Integer, DayPlan> byDate = new LinkedHashMap<>();
        DayPlan birthday = null;

        for (int s = 0; s < HWP.getBodyText().getSectionList().size(); s++) {
            Section sec = HWP.getBodyText().getSectionList().get(s);
            for (int p = 0; p < sec.getParagraphCount(); p++) {
                Paragraph para = sec.getParagraph(p);
                if (para.getControlList() == null) continue;
                for (Control c : para.getControlList()) {
                    if (c.getType() != ControlType.Table) continue;
                    ControlTable ct = (ControlTable) c;

                    // Debug TSV is not needed in AllInOne -> pass null safely
                    DayPlan found = scanTable(ct, byDate, null);
                    if (found != null) birthday = found;
                }
            }
        }

        List<DayPlan> list = new ArrayList<>(byDate.values());
        list.sort(Comparator.comparingInt(dp -> dp.date == null ? Integer.MAX_VALUE : dp.date));

        for (DayPlan dp : list) {
            if (dp.menus == null || dp.menus.isEmpty()) dp.menus = List.of("없음");
            if (dp.pmDesert == null || dp.pmDesert.isEmpty()) dp.pmDesert = List.of("없음");
        }
        if (birthday != null) {
            if (birthday.menus == null || birthday.menus.isEmpty()) birthday.menus = List.of("없음");
            if (birthday.pmDesert == null || birthday.pmDesert.isEmpty()) birthday.pmDesert = List.of("없음");
        }

        List<DayPlan> finalOut = new ArrayList<>(list);
        if (birthday != null) finalOut.add(birthday);

        MAPPER.writeValue(outputJsonPath.toFile(), finalOut);
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

        Path inPath = inputHwpPath.toAbsolutePath();

        int[] ym = parseYearMonthFromFilename(inPath.getFileName().toString());
        if (ym != null) {
            PARSED_YEAR = ym[0];
            PARSED_MONTH = ym[1];
        } else {
            PARSED_YEAR = null;
            PARSED_MONTH = null;
        }

        HWP = HWPReader.fromFile(inPath.toString());

        Map<Integer, DayPlan> byDate = new LinkedHashMap<>();
        DayPlan birthday = null;

        // English comment: Keep the original table scanning logic (same as convert())
        for (int s = 0; s < HWP.getBodyText().getSectionList().size(); s++) {
            Section sec = HWP.getBodyText().getSectionList().get(s);
            for (int p = 0; p < sec.getParagraphCount(); p++) {
                Paragraph para = sec.getParagraph(p);
                if (para.getControlList() == null) continue;
                for (Control c : para.getControlList()) {
                    if (c.getType() != ControlType.Table) continue;
                    ControlTable ct = (ControlTable) c;

                    // Debug TSV is not needed in AllInOne -> pass null safely
                    DayPlan found = scanTable(ct, byDate, null);
                    if (found != null) birthday = found;
                }
            }
        }

        List<DayPlan> list = new ArrayList<>(byDate.values());
        list.sort(Comparator.comparingInt(dp -> dp.date));

        // English comment: Append birthday plan (if any) after normal days (keeps existing behavior)
        if (birthday != null && !list.contains(birthday)) {
            list.add(birthday);
        }

        return new ObjectMapper()
                .enable(SerializationFeature.INDENT_OUTPUT)
                .writeValueAsString(list);
    }
}