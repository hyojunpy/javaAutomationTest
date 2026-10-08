package org.example;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;
import java.nio.file.*;
import java.util.List;
import java.util.stream.Stream;
import static org.junit.jupiter.api.Assertions.*;

/** Eight real meal inputs; opt in with mvn -Pregression verify. */
class ConversionRegressionIT {
    @TempDir Path output;

    static Stream<Path> sources() throws Exception {
        String filter = System.getProperty("conversion.regression.inputFilter", "");
        try (var files = Files.list(Path.of("input"))) {
            return files.filter(p -> p.toString().endsWith(".hwp"))
                    .filter(p -> p.getFileName().toString().contains(filter)).sorted().toList().stream();
        }
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("sources")
    void outputMatchesOriginal(Path input) throws Exception {
        String fixture = "/golden/" + input.getFileName() + ".golden.json";
        try (var stream = getClass().getResourceAsStream(fixture)) {
            assertNotNull(stream, "Missing reviewed original fixture: " + fixture);
            var expected = GoldenFiles.JSON.readTree(stream);
            Path result = new ConversionEngine(Path.of(".")).convert(input, output.resolve("result.xlsx"), ignored -> {});
            var actual = GoldenFiles.capture(input, result);
            assertEquals(expected.get("plan"), actual.get("plan"), "Parsed meal plan changed");
            var expectedParts = expected.get("workbookParts");
            var actualParts = actual.get("workbookParts");
            assertEquals(expectedParts.size(), actualParts.size(), "Workbook part count changed");
            expectedParts.fields().forEachRemaining(part -> assertEquals(part.getValue(), actualParts.get(part.getKey()), "Workbook part changed: " + part.getKey()));
        }
    }
}
