package org.example;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;
import java.util.function.Consumer;

/** UI-independent orchestration; workers own their state and receive resource paths explicitly. */
public final class ConversionEngine {
    private final Path appHome;
    public ConversionEngine(Path home) { appHome = Objects.requireNonNull(home).toAbsolutePath(); }
    public Path convert(Path input, Path output, Consumer<String> status) throws Exception {
        Objects.requireNonNull(input, "input");
        Objects.requireNonNull(output, "output");
        Objects.requireNonNull(status, "status");
        if (!Files.isRegularFile(input)) throw new IllegalArgumentException("HWP 파일을 찾을 수 없습니다: " + input);
        if (!Files.isDirectory(appHome.resolve("input"))) {
            throw new IllegalStateException("input 폴더가 없습니다: " + appHome.resolve("input"));
        }
        boolean extended = SourceMetadata.from(input).extended();
        status.accept("식단 분석 중");
        String json = extended ? HwpToJson.convertToJsonString(input) : HwpToJsonGeneral.convertToJsonString(input);
        status.accept("조리지시서 작성 중 · 일반형은 수 분 걸릴 수 있습니다");
        return extended ? JsonToExcel.convertFromJsonString(json, input, output, appHome)
                : JsonToExcelGeneral.convertFromJsonString(json, input, output, appHome);
    }
}
