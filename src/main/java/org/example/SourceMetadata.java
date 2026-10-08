package org.example;
import java.nio.file.Path;

/** Keep the original, distinct age filename rules of both conversion modes. */
public record SourceMetadata(String filename, boolean extended, boolean age12, boolean generalAge12) {
    public static SourceMetadata from(Path source) {
        String name = source.getFileName().toString();
        return new SourceMetadata(name, name.contains("시간연장"),
                name.contains("만1-2세") || name.contains("만1~2세") || name.contains("만1–2세"),
                name.contains("만1-2"));
    }
    public String description() {
        return (extended ? "시간연장형" : "일반형") + " · "
                + ((extended ? age12 : generalAge12) ? "만1–2세" : "만3–5세");
    }
    public static String outputName(String name) {
        int dot = name.lastIndexOf('.');
        return (dot > 0 ? name.substring(0, dot) : name) + "_수정.xlsx";
    }
}
