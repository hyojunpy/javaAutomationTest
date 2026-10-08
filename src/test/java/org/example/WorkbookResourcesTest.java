package org.example;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import java.nio.file.*;
import static org.junit.jupiter.api.Assertions.*;

class WorkbookResourcesTest {
    @TempDir Path directory;

    @Test void workbookCanBeReplacedAfterReading() throws Exception {
        Path file = directory.resolve("template.xlsx");
        try (var book = new XSSFWorkbook(); var stream = Files.newOutputStream(file)) {
            book.createSheet("template").createRow(0).createCell(0).setCellValue("sample");
            book.write(stream);
        }
        try (var book = WorkbookResources.open(file)) {
            assertEquals("sample", book.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
        }
        Path renamed = Files.move(file, directory.resolve("renamed.xlsx"));
        assertTrue(Files.isRegularFile(renamed));
    }

    @Test void malformedWorkbookDoesNotLeaveInputStreamOpen() throws Exception {
        Path file = Files.writeString(directory.resolve("broken.xlsx"), "invalid workbook");
        assertThrows(Exception.class, () -> WorkbookResources.open(file));
        assertNotNull(Files.move(file, directory.resolve("broken-renamed.xlsx")));
        assertNull(WorkbookResources.open(null));
    }
}
