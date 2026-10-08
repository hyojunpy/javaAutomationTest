package org.example;

import javafx.concurrent.Task;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import java.awt.Desktop;
import java.io.File;
import java.nio.file.Path;
import java.util.function.Consumer;

/** JavaFX adapter for the UI-independent conversion engine. */
public class ConverterService extends Task<Path> {
    private static final Logger LOG = LoggerFactory.getLogger(ConverterService.class);
    private final File hwp;
    private final File output;
    private final Consumer<String> log;
    private final Consumer<String> status;
    private final boolean openFolderAfter;

    public ConverterService(File hwp, File output, Consumer<String> log,
                            Consumer<String> status, boolean openFolderAfter) {
        this.hwp = hwp;
        this.output = output;
        this.log = log;
        this.status = status;
        this.openFolderAfter = openFolderAfter;
    }

    @Override
    protected Path call() throws Exception {
        updateProgress(-1, 1);
        Path home = AppHomeResolver.getAppHome(ConverterService.class);
        log.accept("선택한 식단: " + SourceMetadata.from(hwp.toPath()).description());
        LOG.info("Converting {} to {}", hwp, output);
        Path result = new ConversionEngine(home).convert(hwp.toPath(), output.toPath(), message -> {
            log.accept(message);
            status.accept(message);
        });
        log.accept("저장 완료: " + result.toAbsolutePath());
        if (openFolderAfter) {
            try {
                if (Desktop.isDesktopSupported()) Desktop.getDesktop().open(result.toAbsolutePath().getParent().toFile());
            } catch (Exception exception) {
                LOG.warn("Cannot open result folder", exception);
                log.accept("폴더를 자동으로 열지 못했습니다. 저장 위치를 직접 확인해주세요.");
            }
        }
        updateProgress(1, 1);
        return result;
    }
}
