package org.example;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import java.nio.file.*;
import static org.junit.jupiter.api.Assertions.*;

class ConversionInfrastructureTest {
    @TempDir Path directory;

    @Test void preservesDifferentAgeRules() {
        SourceMetadata alternate = SourceMetadata.from(Path.of("2026년 2월 시간연장형(만1~2세).hwp"));
        assertTrue(alternate.extended()); assertTrue(alternate.age12()); assertFalse(alternate.generalAge12());
        assertEquals("2026년 2월 일반형(만1-2세)_수정.xlsx", SourceMetadata.outputName("2026년 2월 일반형(만1-2세).hwp"));
    }

    @Test void homeOldTemplateTakesPrecedenceOverWorkingDirectoryNew() throws Exception {
        Path home = Files.createDirectories(directory.resolve("home/input"));
        Path working = Files.createDirectories(directory.resolve("cwd/input"));
        Path old = Files.createFile(home.resolve(TemplateCatalog.filename(false, true, false)));
        Files.createFile(working.resolve(TemplateCatalog.filename(false, true, true)));
        TemplateCatalog catalog = new TemplateCatalog(home.getParent(), working.getParent());
        assertEquals(old, catalog.base(false, true));
        Path newer = Files.createFile(home.resolve(TemplateCatalog.filename(false, true, true)));
        assertEquals(newer, catalog.base(false, true));
        assertEquals(old, catalog.older(false, true));
        assertEquals(newer, catalog.newer(false, true));
    }

    @Test void missingResourcesFailWithActionablePath() {
        var failure = assertThrows(IllegalStateException.class, () -> new TemplateCatalog(directory, directory).base(true, false));
        assertTrue(failure.getMessage().contains("input"));
    }

    @Test void engineRejectsMissingInputBeforeRendering() {
        assertThrows(IllegalArgumentException.class, () -> new ConversionEngine(directory).convert(directory.resolve("missing.hwp"), directory.resolve("output.xlsx"), ignored -> {}));
    }

    @Test void generalRendererKeepsNullInputErrorContract() {
        var error = assertThrows(IllegalArgumentException.class,
                () -> JsonToExcelGeneral.convertFromJsonString("[]", null, directory.resolve("output.xlsx")));
        assertEquals("sourceNamePath is null", error.getMessage());
    }

    @Test void parserCallsDoNotShareFilenameState() throws Exception {
        Path source = Path.of("input/2026년 2월 시간연장형(만1-2세).hwp");
        String expected = HwpToJson.convertToJsonString(source);
        var executor = java.util.concurrent.Executors.newFixedThreadPool(2);
        try {
            var one = executor.submit(() -> HwpToJson.convertToJsonString(source));
            var two = executor.submit(() -> HwpToJson.convertToJsonString(Path.of("input/2026년 1월 시간연장형(만3-5세).hwp")));
            assertEquals(expected, one.get()); assertNotNull(two.get());
        } finally { executor.shutdownNow(); }
        assertEquals(expected, HwpToJson.convertToJsonString(source));
    }
}
