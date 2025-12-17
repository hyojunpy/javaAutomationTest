package org.example;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class AllInOne {

    public static void main(String[] args) throws Exception {

        // If no args -> launch UI
        if (args == null || args.length == 0) {
            AllInOneUI.main(new String[0]);
            return;
        }

        Path inputHwp = Paths.get(args[0]).toAbsolutePath();

        if (!Files.exists(inputHwp)) {
            System.out.println("Input HWP not found: " + inputHwp);
            System.exit(1);
        }

        Path outXlsx;
        if (args.length >= 2 && args[1] != null && args[1].trim().length() > 0) {
            // If user provides output path
            Path outArg = Paths.get(args[1].trim()).toAbsolutePath();

            // If it's an existing directory -> auto filename inside it
            if (Files.exists(outArg) && Files.isDirectory(outArg)) {
                outXlsx = outArg.resolve(makeOutName(inputHwp.getFileName().toString()));
            } else {
                // Otherwise treat as file path
                outXlsx = outArg;
            }
        } else {
            // Auto output path: <user home>/Desktop/<inputName>_수정.xlsx (safer than Program Files)
            Path desktop = Paths.get(System.getProperty("user.home"), "Desktop").toAbsolutePath();
            Files.createDirectories(desktop);
            outXlsx = desktop.resolve(makeOutName(inputHwp.getFileName().toString()));
        }

        // Ensure output directory exists
        if (outXlsx.getParent() != null) {
            Files.createDirectories(outXlsx.getParent());
        }

        // Decide pipeline by filename
        String name = inputHwp.getFileName().toString();
        boolean isGeneral = name.contains("일반형");

        // IMPORTANT: Place intermediate JSON next to the output XLSX (user-selected output folder)
        Path jsonPath = buildJsonPathFromInput(inputHwp, outXlsx);

        System.out.println("[INPUT ] " + inputHwp);
        System.out.println("[MID   ] " + jsonPath);
        System.out.println("[OUTPUT] " + outXlsx);
        System.out.println("[MODE  ] " + (isGeneral ? "GENERAL" : "EXTENDED"));

        if (isGeneral) {
            HwpToJsonGeneral.main(new String[]{ inputHwp.toString(), jsonPath.toString() });
            JsonToExcelGeneral.convert(jsonPath, outXlsx);
        } else {
            HwpToJson.main(new String[]{ inputHwp.toString(), jsonPath.toString() });
            JsonToExcel.convert(jsonPath, outXlsx);
        }

        System.out.println("DONE");
    }

    private static String makeOutName(String inputFileName) {
        String stem = inputFileName;
        int dot = stem.lastIndexOf('.');
        if (dot > 0) {
            stem = stem.substring(0, dot);
        }
        return stem + "_수정.xlsx";
    }

    private static Path buildJsonPathFromInput(Path inputHwp, Path outXlsx) throws Exception {
        // Use the output XLSX directory for JSON to avoid write-permission issues in Program Files
        Path outDir = outXlsx.toAbsolutePath().getParent();
        if (outDir == null) {
            throw new IllegalArgumentException("Output directory is null: " + outXlsx);
        }
        Files.createDirectories(outDir);

        String stem = inputHwp.getFileName().toString();
        int dot = stem.lastIndexOf('.');
        if (dot > 0) {
            stem = stem.substring(0, dot);
        }

        return outDir.resolve(stem + ".json");
    }
}
