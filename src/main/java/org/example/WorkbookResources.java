package org.example;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

final class WorkbookResources {
    private WorkbookResources() {}
    static Workbook open(Path path) throws Exception {
        if (path == null) return null;
        try (InputStream input = Files.newInputStream(path)) {
            return WorkbookFactory.create(input);
        }
    }
}
