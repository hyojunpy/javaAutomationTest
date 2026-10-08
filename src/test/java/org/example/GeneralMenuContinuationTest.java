package org.example;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;
import java.nio.file.Path;
import java.util.List;
import java.util.stream.Stream;
import static org.junit.jupiter.api.Assertions.*;

class GeneralMenuContinuationTest {
    private final HwpToJsonGeneral parser = new HwpToJsonGeneral();
    private static final String ALLERGENS = "①②⑤⑥⑩⑫⑮⑯⑱";
    private static final String SAUCE = "&소스①⑤⑥⑩⑫⑮⑯";

    @Test void joinsWrappedAllergensAndSauceKeepingAllergenNumbers() {
        assertEquals(List.of("함박스테이크" + ALLERGENS + SAUCE, "쌀밥½"),
                parser.splitMenusByLineThenSlash("함박스테이크\n" + ALLERGENS + SAUCE + "\n쌀밥½"));
    }
    @Test void keepsExistingAmpersandContinuation() {
        assertEquals(List.of("함박스테이크" + ALLERGENS + SAUCE),
                parser.splitMenusByLineThenSlash("함박스테이크" + ALLERGENS + "\n" + SAUCE));
    }
    @Test void joinsThreeWrappedLinesFromActualHwp() {
        assertEquals(List.of("함박스테이크" + ALLERGENS + SAUCE, "쌀밥½", "깍두기⑨"),
                parser.splitMenusByLineThenSlash("함박스테이크\n" + ALLERGENS + "\n" + SAUCE + "\n쌀밥½/깍두기⑨"));
    }
    @Test void acceptsCrLfAndSpacesAroundContinuation() {
        assertEquals(List.of("함박스테이크①② &소스⑤"),
                parser.splitMenusByLineThenSlash("함박스테이크\r\n ①② &소스⑤"));
    }
    @Test void doesNotConsumeContinuationWithoutPreviousMenu() {
        assertEquals(List.of(ALLERGENS + SAUCE), parser.splitMenusByLineThenSlash(ALLERGENS + SAUCE));
    }
    @Test void doesNotJoinAcrossCells() {
        assertEquals(List.of("함박스테이크"), parser.splitMenusByLineThenSlash("함박스테이크"));
        assertEquals(List.of(ALLERGENS + SAUCE), parser.splitMenusByLineThenSlash(ALLERGENS + SAUCE));
    }
    @Test void doesNotJoinOtherMenuOrStandaloneAllergens() {
        assertEquals(List.of("함박스테이크", "①②", "①②소스⑤", "탕수육⑤⑩&소스"),
                parser.splitMenusByLineThenSlash("함박스테이크\n①②\n①②소스⑤\n탕수육⑤⑩&소스"));
    }
    @Test void doesNotJoinNewPatternAcrossSlashSeparator() {
        assertEquals(List.of("함박스테이크", ALLERGENS + SAUCE),
                parser.splitMenusByLineThenSlash("함박스테이크/" + ALLERGENS + SAUCE));
    }
    @Test void joinsOnlyWrappedFirstPartBeforeSlash() {
        assertEquals(List.of("함박스테이크" + ALLERGENS + SAUCE, "쌀밥½"),
                parser.splitMenusByLineThenSlash("함박스테이크\n" + ALLERGENS + SAUCE + "/쌀밥½"));
    }
    @Test void doesNotFoldSlashSeparatedAllergensIntoPreviousMenu() {
        assertEquals(List.of("함박스테이크", ALLERGENS + SAUCE),
                parser.splitMenusByLineThenSlash("함박스테이크/" + ALLERGENS + "\n" + SAUCE));
    }

    static Stream<Path> mealSources() throws Exception { return ConversionRegressionIT.sources(); }

    @ParameterizedTest(name = "meal plan: {0}")
    @MethodSource("mealSources")
    void realMealPlansMatchReviewedExpectations(Path source) throws Exception {
        try (var input = getClass().getResourceAsStream("/golden/" + source.getFileName() + ".golden.json")) {
            assertNotNull(input);
            var expected = GoldenFiles.JSON.readTree(input).get("plan");
            String parsed = SourceMetadata.from(source).extended()
                    ? HwpToJson.convertToJsonString(source) : HwpToJsonGeneral.convertToJsonString(source);
            assertEquals(expected, GoldenFiles.JSON.readTree(parsed));
        }
    }
}
