package org.example;

import javafx.application.Platform;
import javafx.concurrent.Task;

import java.awt.Desktop;
import java.io.File;
import java.nio.file.Path;
import java.util.function.Consumer;

/**
 * Background conversion task.
 */
public class ConverterService extends Task<Void> {

    private final File hwp;
    private final File outXlsx;
    private final Consumer<String> log;
    private final Consumer<String> status;
    private final boolean autoExpandLog;
    private final boolean openFolderAfter;

    public ConverterService(File hwp, File outXlsx, Consumer<String> log, Consumer<String> status, boolean autoExpandLog, boolean openFolderAfter) {
        this.hwp = hwp;
        this.outXlsx = outXlsx;
        this.log = log;
        this.status = status;
        this.autoExpandLog = autoExpandLog;
        this.openFolderAfter = openFolderAfter;
    }

    @Override
    protected Void call() throws Exception {
        updateProgress(-1, 1);

        status.accept("appHome 확인 중...");
        Path appHome = AppHomeResolver.getAppHome(ConverterService.class);
        log.accept("appHome: " + appHome);
        log.accept("inputDir: " + appHome.resolve("input"));

        File inputDir = appHome.resolve("input").toFile();
        if (!inputDir.exists() || !inputDir.isDirectory()) {
            throw new IllegalStateException("input 폴더가 없습니다: " + inputDir.getAbsolutePath());
        }

        // English comment: Many existing codes use "app.home" system property to resolve input resources.
        System.setProperty("app.home", appHome.toString());

        boolean isExtended = hwp.getName().contains("시간연장");
        log.accept("MODE: " + (isExtended ? "EXTENDED" : "GENERAL"));

        status.accept("HWP 분석 중...");
        String json;

        if (isExtended) {
            json = HwpToJson.convertToJsonString(hwp);
            status.accept("엑셀 생성 중...");
            JsonToExcel.convertFromJsonString(json, hwp.toPath(), outXlsx.toPath());
        } else {
            json = HwpToJsonGeneral.convertToJsonString(hwp);
            status.accept("엑셀 생성 중...");
            JsonToExcelGeneral.convertFromJsonString(json, hwp.toPath(), outXlsx.toPath());
        }

        log.accept("OUTPUT: " + outXlsx.getAbsolutePath());

        if (openFolderAfter) {
            try {
                if (Desktop.isDesktopSupported()) {
                    Desktop.getDesktop().open(outXlsx.getParentFile());
                }
            } catch (Exception ignore) {
                // English comment: Ignore folder open errors.
            }
        }

        updateProgress(1, 1);
        return null;
    }
}
