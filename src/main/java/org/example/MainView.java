package org.example;

import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.application.Platform;
import javafx.beans.property.BooleanProperty;
import javafx.beans.property.SimpleBooleanProperty;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.input.TransferMode;
import javafx.scene.layout.*;
import javafx.stage.FileChooser;
import javafx.util.Duration;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import java.awt.Desktop;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Locale;

/** File selection, conversion status and result actions. Domain rules live in ConversionEngine. */
public class MainView {
    private static final Logger LOG = LoggerFactory.getLogger(MainView.class);
    private final BorderPane root = new BorderPane();
    private final TextField inputField = new TextField();
    private final TextField outputField = new TextField();
    private final Label selectedName = new Label("아직 선택한 식단이 없습니다");
    private final Label metadata = new Label("일반형 · 시간연장형 / 만1–2세 · 만3–5세");
    private final Label statusLabel = new Label("식단 파일을 선택하면 준비가 완료됩니다");
    private final Label elapsedLabel = new Label("");
    private final Button pick = new Button("HWP 파일 선택");
    private final Button changeOutput = new Button("저장 위치 변경");
    private final Button convert = new Button("조리지시서 만들기");
    private final Button openResult = new Button("Excel 열기");
    private final Button openFolder = new Button("저장 폴더 열기");
    private final CheckBox autoOpen = new CheckBox("완료 후 저장 폴더 열기");
    private final CheckBox autoLog = new CheckBox("작업 내역 자동 펼치기");
    private final ProgressBar progress = new ProgressBar(0);
    private final TextArea logs = new TextArea();
    private final TitledPane logPane = new TitledPane();
    private final BooleanProperty running = new SimpleBooleanProperty();
    private final Timeline clock = new Timeline(new KeyFrame(Duration.seconds(1), e -> updateElapsed()));
    private long startedAt;
    private Path lastResult;
    private boolean outputChosen;

    public MainView() { build(); wire(); }
    public Parent getRoot() { return root; }
    public void applyStyles(Scene scene) {
        String stylesheet = getClass().getResource("/styles.css").toExternalForm();
        if (!scene.getStylesheets().contains(stylesheet)) scene.getStylesheets().add(stylesheet);
    }

    private void build() {
        root.setPadding(new Insets(28));
        root.setId("main-view");
        Label eyebrow = new Label("MEAL WORKSPACE");
        eyebrow.getStyleClass().add("eyebrow");
        Label title = new Label("조리지시서 변환기");
        title.getStyleClass().add("title");
        Label subtitle = new Label("매월 반복되는 조리지시서 작성, 식단 파일 하나로 시작하세요.");
        subtitle.getStyleClass().add("muted");
        VBox header = new VBox(7, eyebrow, title, subtitle);
        header.setPadding(new Insets(0, 0, 22, 0));
        root.setTop(header);

        Label dropTitle = new Label("식단 HWP를 여기에 끌어다 놓으세요");
        dropTitle.getStyleClass().add("drop-title");
        Label hint = new Label("HWP 파일 1개를 선택해주세요. 유형과 연령은 파일명으로 확인합니다.");
        hint.getStyleClass().add("muted");
        hint.setWrapText(true);
        pick.getStyleClass().add("secondary");
        pick.setId("pick-input");
        VBox drop = new VBox(9, dropTitle, hint, pick);
        drop.setAlignment(Pos.CENTER);
        drop.setPadding(new Insets(23));
        drop.getStyleClass().add("drop-zone");
        drop.setId("drop-zone");

        selectedName.getStyleClass().add("section-title");
        selectedName.setWrapText(true);
        metadata.getStyleClass().add("badge");
        metadata.setId("source-metadata");
        inputField.setEditable(false);
        outputField.setEditable(false);
        inputField.setId("input-path");
        outputField.setId("output-path");
        inputField.setPromptText("선택한 HWP 파일 경로");
        outputField.setPromptText("Excel 저장 위치");
        Tooltip inputTip = new Tooltip(); inputTip.textProperty().bind(inputField.textProperty()); inputField.setTooltip(inputTip);
        Tooltip outputTip = new Tooltip(); outputTip.textProperty().bind(outputField.textProperty()); outputField.setTooltip(outputTip);
        HBox outputRow = new HBox(10, outputField, changeOutput);
        HBox.setHgrow(outputField, Priority.ALWAYS);
        Label outputTitle = new Label("저장할 Excel 파일");
        outputTitle.getStyleClass().add("field-label");
        VBox fileCard = new VBox(10, selectedName, metadata, inputField, outputTitle, outputRow);
        fileCard.getStyleClass().add("card");

        HBox options = new HBox(20, autoOpen, autoLog);
        options.setAlignment(Pos.CENTER_LEFT);
        progress.setMaxWidth(Double.MAX_VALUE);
        statusLabel.setWrapText(true);
        statusLabel.setId("conversion-status");
        statusLabel.getStyleClass().add("status-text");
        Region spacer = new Region(); HBox.setHgrow(spacer, Priority.ALWAYS);
        HBox state = new HBox(12, statusLabel, spacer, elapsedLabel);
        state.setAlignment(Pos.CENTER_LEFT);
        elapsedLabel.getStyleClass().add("muted");
        VBox statusCard = new VBox(12, state, progress);
        statusCard.getStyleClass().add("status-card");

        openResult.setDisable(true); openFolder.setDisable(true);
        openResult.setId("open-result"); openFolder.setId("open-folder");
        convert.setId("convert"); convert.getStyleClass().add("primary");
        convert.setDefaultButton(true);
        Region actionsSpace = new Region(); HBox.setHgrow(actionsSpace, Priority.ALWAYS);
        HBox actions = new HBox(10, openResult, openFolder, actionsSpace, convert);
        actions.setAlignment(Pos.CENTER_LEFT);
        logs.setEditable(false); logs.setWrapText(true); logs.setPrefRowCount(6);
        logPane.setText("작업 내역"); logPane.setExpanded(false); logPane.setContent(logs);
        VBox content = new VBox(16, drop, fileCard, options, statusCard, actions, logPane);
        ScrollPane scroll = new ScrollPane(content);
        scroll.setFitToWidth(true); scroll.getStyleClass().add("workspace-scroll");
        root.setCenter(scroll);
        Label footer = new Label("파일명 예시: 2026년 2월 일반형(만1-2세).hwp  ·  생성 후 메뉴와 수량을 확인해주세요.");
        footer.setWrapText(true); footer.getStyleClass().add("footer");
        BorderPane.setMargin(footer, new Insets(16, 0, 0, 0)); root.setBottom(footer);

        pick.disableProperty().bind(running);
        changeOutput.disableProperty().bind(running.or(inputField.textProperty().isEmpty()));
        autoOpen.disableProperty().bind(running); autoLog.disableProperty().bind(running);
        convert.disableProperty().bind(running.or(inputField.textProperty().isEmpty()).or(outputField.textProperty().isEmpty()));
        clock.setCycleCount(Timeline.INDEFINITE);
    }

    private void wire() {
        root.setOnDragOver(e -> {
            if (!running.get() && e.getDragboard().hasFiles() && e.getDragboard().getFiles().size() == 1
                    && isHwp(e.getDragboard().getFiles().get(0))) e.acceptTransferModes(TransferMode.COPY);
            e.consume();
        });
        root.setOnDragDropped(e -> {
            boolean accepted = !running.get() && e.getDragboard().hasFiles()
                    && e.getDragboard().getFiles().size() == 1 && isHwp(e.getDragboard().getFiles().get(0));
            if (accepted) selectInput(e.getDragboard().getFiles().get(0).toPath());
            e.setDropCompleted(accepted); e.consume();
        });
        pick.setOnAction(e -> {
            FileChooser chooser = new FileChooser(); chooser.setTitle("식단 HWP 선택");
            chooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("HWP 식단 파일", "*.hwp", "*.HWP"));
            File file = chooser.showOpenDialog(root.getScene().getWindow());
            if (file != null) selectInput(file.toPath());
        });
        changeOutput.setOnAction(e -> chooseOutput());
        convert.setOnAction(e -> runConvert());
        openResult.setOnAction(e -> open(lastResult));
        openFolder.setOnAction(e -> open(lastResult == null ? null : lastResult.getParent()));
    }

    void selectInput(Path source) {
        if (running.get()) return;
        if (!isHwp(source.toFile()) || !Files.isRegularFile(source)) {
            showStatus("선택한 HWP 파일을 찾을 수 없습니다.", "error"); return;
        }
        source = source.toAbsolutePath();
        inputField.setText(source.toString());
        selectedName.setText(source.getFileName().toString());
        metadata.setText(SourceMetadata.from(source).description());
        Path parent = outputChosen ? Path.of(outputField.getText()).getParent() : source.getParent();
        if (parent != null) outputField.setText(parent.resolve(SourceMetadata.outputName(source.getFileName().toString())).toString());
        resetResult(); elapsedLabel.setText(""); progress.setProgress(0);
        showStatus("준비 완료 · 저장 위치를 확인하고 변환을 시작하세요.", "ready");
        log("입력: " + source);
    }

    private void chooseOutput() {
        FileChooser chooser = new FileChooser(); chooser.setTitle("조리지시서 저장 위치");
        chooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel 파일", "*.xlsx"));
        File current = new File(outputField.getText());
        if (current.getParentFile() != null && current.getParentFile().isDirectory()) chooser.setInitialDirectory(current.getParentFile());
        chooser.setInitialFileName(current.getName());
        File file = chooser.showSaveDialog(root.getScene().getWindow());
        if (file != null) {
            String path = file.getAbsolutePath();
            outputField.setText(path.toLowerCase(Locale.ROOT).endsWith(".xlsx") ? path : path + ".xlsx");
            outputChosen = true; resetResult();
        }
    }

    private void runConvert() {
        if (running.get()) return;
        File input = new File(inputField.getText()), output = new File(outputField.getText());
        if (!input.isFile()) { showStatus("입력 파일이 없습니다. 다시 선택해주세요.", "error"); return; }
        if (output.exists()) {
            Alert alert = new Alert(Alert.AlertType.CONFIRMATION, "같은 이름의 Excel 파일이 있습니다. 덮어쓸까요?", ButtonType.OK, ButtonType.CANCEL);
            alert.initOwner(root.getScene().getWindow()); alert.setHeaderText("기존 파일 덮어쓰기");
            if (alert.showAndWait().orElse(ButtonType.CANCEL) != ButtonType.OK) return;
        }
        resetResult(); running.set(true); startedAt = System.nanoTime();
        updateElapsed(); clock.playFromStart();
        if (autoLog.isSelected()) logPane.setExpanded(true);
        showStatus("변환 준비 중", "working");
        ConverterService task = new ConverterService(input, output, this::log,
                msg -> Platform.runLater(() -> showStatus(msg, "working")), autoOpen.isSelected());
        progress.progressProperty().bind(task.progressProperty());
        task.setOnSucceeded(e -> {
            finish(); lastResult = task.getValue().toAbsolutePath(); progress.setProgress(1);
            openResult.setDisable(false); openFolder.setDisable(false);
            showStatus("변환 완료 · Excel에서 메뉴와 수량을 확인해주세요.", "success");
        });
        task.setOnFailed(e -> {
            finish(); progress.setProgress(0); logPane.setExpanded(true);
            Throwable failure = task.getException();
            String reason = failure == null ? "원인을 확인하지 못했습니다." : failure.getMessage();
            showStatus("변환 실패 · " + reason, "error"); log("오류: " + reason);
            LOG.error("Conversion failed for {}", input, failure);
            if (failure != null) {
                java.io.StringWriter trace = new java.io.StringWriter();
                failure.printStackTrace(new java.io.PrintWriter(trace));
                log(trace.toString());
            }
        });
        Thread worker = new Thread(task, "meal-conversion"); worker.setDaemon(true); worker.start();
    }

    private void finish() { clock.stop(); updateElapsed(); progress.progressProperty().unbind(); running.set(false); }
    private void updateElapsed() {
        long seconds = (System.nanoTime() - startedAt) / 1_000_000_000L;
        elapsedLabel.setText(String.format("%02d:%02d", seconds / 60, seconds % 60));
    }
    private void resetResult() { lastResult = null; openResult.setDisable(true); openFolder.setDisable(true); }
    private void showStatus(String message, String state) {
        statusLabel.setText(message);
        statusLabel.getStyleClass().removeAll("ready", "working", "success", "error");
        statusLabel.getStyleClass().add(state);
    }
    private boolean isHwp(File file) { return file != null && file.getName().toLowerCase(Locale.ROOT).endsWith(".hwp"); }
    private void log(String message) {
        Platform.runLater(() -> { logs.appendText(message + "\n"); logs.positionCaret(logs.getLength()); });
    }
    private void open(Path path) {
        if (path == null) return;
        try {
            if (!Desktop.isDesktopSupported()) throw new UnsupportedOperationException("파일 열기를 지원하지 않습니다.");
            Desktop.getDesktop().open(path.toFile());
        } catch (Exception exception) {
            LOG.warn("Cannot open {}", path, exception);
            logPane.setExpanded(true); log("파일을 열지 못했습니다: " + exception.getMessage());
        }
    }
}
