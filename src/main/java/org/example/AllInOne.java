package org.example;

import javax.swing.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class AllInOne {

    public static void main(String[] args) throws Exception {

        // If no args -> launch UI
        if (args == null || args.length == 0) {
            setupLookAndFeel();
            javax.swing.SwingUtilities.invokeLater(() -> new AllInOneUI().setVisible(true));
            return;
        }

        Path inputHwp = Paths.get(args[0]).toAbsolutePath();

        if (!Files.exists(inputHwp)) {
            System.out.println("Input HWP not found: " + inputHwp);
            System.exit(1);
        }

        Path outXlsx;
        if (args.length >= 2 && args[1] != null && args[1].trim().length() > 0) {
            Path outArg = Paths.get(args[1].trim()).toAbsolutePath();
            if (Files.exists(outArg) && Files.isDirectory(outArg)) {
                outXlsx = outArg.resolve(makeOutName(inputHwp.getFileName().toString()));
            } else {
                outXlsx = outArg;
            }
        } else {
            Path desktop = Paths.get(System.getProperty("user.home"), "Desktop").toAbsolutePath();
            Files.createDirectories(desktop);
            outXlsx = desktop.resolve(makeOutName(inputHwp.getFileName().toString()));
        }

        if (outXlsx.getParent() != null) {
            Files.createDirectories(outXlsx.getParent());
        }

        // ✅ 판별 기준: 파일명에 "시간연장" 포함 여부
        boolean isExtended = inputHwp.getFileName().toString().contains("시간연장");

        System.out.println("[INPUT ] " + inputHwp);
        System.out.println("[OUTPUT] " + outXlsx);
        System.out.println("[MODE  ] " + (isExtended ? "EXTENDED" : "GENERAL"));

        if (isExtended) {
            String json = HwpToJson.convertToJsonString(inputHwp.toFile());
            JsonToExcel.convertFromJsonString(json, inputHwp, outXlsx);
        } else {
            String json = HwpToJsonGeneral.convertToJsonString(inputHwp.toFile());
            JsonToExcelGeneral.convertFromJsonString(json, inputHwp, outXlsx);
        }

        System.out.println("DONE");
    }

    private static void setupLookAndFeel() {
        try {
            UIManager.setLookAndFeel("javax.swing.plaf.nimbus.NimbusLookAndFeel");
        } catch (Exception ignore) {
        }
    }

    private static String makeOutName(String inputFileName) {
        String stem = inputFileName;
        int dot = stem.lastIndexOf('.');
        if (dot > 0) {
            stem = stem.substring(0, dot);
        }
        return stem + "_수정.xlsx";
    }
}
