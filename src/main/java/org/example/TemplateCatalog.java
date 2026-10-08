package org.example;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/** Preserve home/new → home/old → cwd/new → cwd/old lookup order. */
final class TemplateCatalog {
    private final List<Path> directories;
    TemplateCatalog(Path home) { this(home, Path.of("").toAbsolutePath()); }
    TemplateCatalog(Path home, Path workingDirectory) {
        directories = List.of(home.resolve("input"), workingDirectory.resolve("input"));
    }
    static Path defaultHome() {
        String home = System.getProperty("app.home");
        return home == null || home.isBlank() ? Path.of("").toAbsolutePath() : Path.of(home).toAbsolutePath();
    }
    static String filename(boolean extended, boolean age12, boolean newer) {
        String prefix = newer ? "★2026~조리지시서" : (extended ? "2021.9~2025 조리지시서" : "2022~2025 조리지시서");
        return prefix + "(" + (age12 ? "만1-2세" : "만3-5세") + " " + (extended ? "시간연장형" : "일반형") + ").xlsx";
    }
    Path base(boolean extended, boolean age12) {
        for (Path dir : directories) {
            for (boolean newer : new boolean[]{true, false}) {
                Path candidate = dir.resolve(filename(extended, age12, newer));
                if (Files.exists(candidate)) return candidate;
            }
        }
        throw new IllegalStateException("조리지시서 양식을 찾을 수 없습니다: "
                + directories.get(0).resolve(filename(extended, age12, true)));
    }
    Path newer(boolean extended, boolean age12) { return lookup(extended, age12, true); }
    Path older(boolean extended, boolean age12) { return lookup(extended, age12, false); }
    private Path lookup(boolean extended, boolean age12, boolean newer) {
        for (Path dir : directories) {
            Path candidate = dir.resolve(filename(extended, age12, newer));
            if (Files.exists(candidate)) return candidate;
        }
        return null;
    }
}
