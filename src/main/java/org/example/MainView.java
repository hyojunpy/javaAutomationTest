package org.example;

import javafx.application.Platform;
import javafx.beans.binding.Bindings;
import javafx.concurrent.Task;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.input.Dragboard;
import javafx.scene.input.TransferMode;
import javafx.scene.layout.*;
import javafx.stage.FileChooser;

import javafx.beans.property.BooleanProperty;
import javafx.beans.property.SimpleBooleanProperty;


import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * Main JavaFX UI.
 */
public class MainView {

    private final BorderPane root = new BorderPane();

    private final TextField hwpField = new TextField();
    private final TextField outField = new TextField();

    private final Button btnPickHwp = new Button("파일 선택");
    private final Button btnChangeHwp = new Button("변경");
    private final Button btnChangeOut = new Button("변경");

    private final Button btnConvert = new Button("변환 시작");
    private final Button btnExit = new Button("종료");

    private final ProgressBar progress = new ProgressBar(0);
    private final Label statusLabel = new Label("준비됨");

    private final TextArea logArea = new TextArea();
    private final TitledPane logPane = new TitledPane();

    private final CheckBox chkOpenFolder = new CheckBox("변환 후 폴더 열기");
    private final CheckBox chkAutoExpandLog = new CheckBox("로그 자동 펼치기");

    private final BooleanProperty runningProperty = new SimpleBooleanProperty(false);

    private boolean outputChosenByUser = false;

    public MainView() {
        build();
        wireEvents();
    }

    public Parent getRoot() {
        return root;
    }

    public void applyStyles(Scene scene) {
        // English comment: Load CSS from classpath if you add it (see below).
        // scene.getStylesheets().add(getClass().getResource("/styles.css").toExternalForm());
    }

    private void build() {
        root.setPadding(new Insets(14));

        // Header
        Label title = new Label("조리지시서 변환기");
        title.getStyleClass().add("title");

        Label subtitle = new Label("HWP → XLSX  (시간연장형/일반형 자동 판별)");
        subtitle.getStyleClass().add("subtitle");

        VBox header = new VBox(6, title, subtitle);
        header.setPadding(new Insets(2, 2, 12, 2));

        // Drop zone
        VBox dropZone = createDropZone();

        // Paths area
        VBox pathBox = new VBox(10,
                createPathRow("입력 HWP", hwpField, btnChangeHwp),
                createPathRow("출력 XLSX", outField, btnChangeOut)
        );

        // Options
        HBox options = new HBox(16, chkOpenFolder, chkAutoExpandLog);
        options.setPadding(new Insets(6, 0, 6, 0));

        // Status + progress
        progress.setPrefWidth(520);
        progress.setMaxWidth(Double.MAX_VALUE);

        VBox statusBox = new VBox(6, statusLabel, progress);
        statusBox.setPadding(new Insets(6, 0, 6, 0));

        // Buttons
        btnConvert.getStyleClass().add("primary");
        HBox buttons = new HBox(10, btnConvert, btnExit);
        buttons.setAlignment(Pos.CENTER_RIGHT);

        // Log
        logArea.setEditable(false);
        logArea.setWrapText(true);
        logArea.setPrefRowCount(10);

        logPane.setText("자세히 보기(로그)");
        logPane.setExpanded(false);
        logPane.setContent(new VBox(logArea));
        logPane.setCollapsible(true);

        VBox center = new VBox(12, dropZone, pathBox, options, statusBox, buttons, logPane);
        VBox.setVgrow(logPane, Priority.ALWAYS);

        root.setTop(header);
        root.setCenter(center);

        // Defaults
        hwpField.setEditable(false);
        outField.setEditable(false);
        btnChangeHwp.setDisable(true);
        btnChangeOut.setDisable(true);

        btnConvert.disableProperty().bind(
                Bindings.createBooleanBinding(
                        () -> hwpField.getText() == null || hwpField.getText().trim().isEmpty() || runningProperty.get(),
                        hwpField.textProperty(),
                        runningProperty
                )
        );
    }

    private VBox createDropZone() {
        Label dropTitle = new Label("여기에 HWP 파일을 끌어다 놓으세요");
        dropTitle.getStyleClass().add("drop-title");

        Label dropHint = new Label("또는 아래 버튼으로 파일을 선택하세요");
        dropHint.getStyleClass().add("drop-hint");

        btnPickHwp.getStyleClass().add("secondary");

        VBox box = new VBox(6, dropTitle, dropHint, btnPickHwp);
        box.setAlignment(Pos.CENTER);
        box.setPadding(new Insets(22));
        box.getStyleClass().add("drop-zone");
        return box;
    }

    private HBox createPathRow(String label, TextField field, Button btnChange) {
        Label lb = new Label(label);
        lb.setMinWidth(90);

        HBox.setHgrow(field, Priority.ALWAYS);
        field.setPrefHeight(30);

        btnChange.setPrefHeight(30);

        HBox row = new HBox(10, lb, field, btnChange);
        row.setAlignment(Pos.CENTER_LEFT);
        return row;
    }

    private void wireEvents() {
        // Drag & Drop
        root.setOnDragOver(e -> {
            Dragboard db = e.getDragboard();
            if (db.hasFiles() && db.getFiles().size() > 0) {
                File f = db.getFiles().get(0);
                if (isHwpFile(f)) {
                    e.acceptTransferModes(TransferMode.COPY);
                }
            }
            e.consume();
        });

        root.setOnDragDropped(e -> {
            Dragboard db = e.getDragboard();
            boolean success = false;
            if (db.hasFiles() && db.getFiles().size() > 0) {
                File f = db.getFiles().get(0);
                if (isHwpFile(f)) {
                    applySelectedHwp(f);
                    success = true;
                }
            }
            e.setDropCompleted(success);
            e.consume();
        });

        btnPickHwp.setOnAction(e -> chooseHwp());
        btnChangeHwp.setOnAction(e -> chooseHwp());
        btnChangeOut.setOnAction(e -> chooseOutputXlsx());

        btnExit.setOnAction(e -> Platform.exit());

        btnConvert.setOnAction(e -> runConvert());
    }

    private void chooseHwp() {
        FileChooser fc = new FileChooser();
        fc.setTitle("HWP 파일 선택");
        fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("HWP Files", "*.hwp"));

        File picked = fc.showOpenDialog(root.getScene().getWindow());
        if (picked != null) {
            applySelectedHwp(picked);
        }
    }

    private void applySelectedHwp(File f) {
        hwpField.setText(f.getAbsolutePath());
        btnChangeHwp.setDisable(false);
        btnChangeOut.setDisable(false);

        // English comment: Auto-suggest output unless user explicitly chose output path.
        if (!outputChosenByUser) {
            File suggested = new File(f.getParentFile(), makeOutName(f.getName()));
            outField.setText(suggested.getAbsolutePath());
        } else {
            File current = new File(outField.getText().trim());
            File parent = current.getParentFile();
            if (parent != null) {
                File suggested = new File(parent, makeOutName(f.getName()));
                outField.setText(suggested.getAbsolutePath());
            }
        }

        log("INPUT: " + f.getAbsolutePath());
    }

    private void chooseOutputXlsx() {
        FileChooser fc = new FileChooser();
        fc.setTitle("출력 XLSX 저장 위치 선택");
        fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Workbook", "*.xlsx"));

        String defaultName = "output.xlsx";
        if (hwpField.getText() != null && hwpField.getText().trim().length() > 0) {
            File hwp = new File(hwpField.getText().trim());
            defaultName = makeOutName(hwp.getName());
            if (hwp.getParentFile() != null) {
                fc.setInitialDirectory(hwp.getParentFile());
            }
        }

        fc.setInitialFileName(defaultName);

        File picked = fc.showSaveDialog(root.getScene().getWindow());
        if (picked != null) {
            String path = picked.getAbsolutePath();
            if (!path.toLowerCase().endsWith(".xlsx")) {
                picked = new File(path + ".xlsx");
            }

            outField.setText(picked.getAbsolutePath());
            outputChosenByUser = true;

            log("OUTPUT: " + picked.getAbsolutePath());
        }
    }

    private void runConvert() {
        String in = hwpField.getText() == null ? "" : hwpField.getText().trim();
        String out = outField.getText() == null ? "" : outField.getText().trim();

        if (in.isEmpty()) {
            status("HWP 파일을 먼저 선택해주세요.");
            return;
        }
        if (out.isEmpty()) {
            status("출력 XLSX 저장 위치를 선택해주세요.");
            return;
        }

        File hwp = new File(in);
        File outXlsx = new File(out);

        if (!hwp.exists() || !hwp.isFile()) {
            status("HWP 파일이 없습니다: " + hwp.getAbsolutePath());
            return;
        }

        // English comment: Disable controls while running.
        setUiRunning(true);
        status("변환 중...");

        Task<Void> task = new ConverterService(hwp, outXlsx, this::log, this::status, chkAutoExpandLog.isSelected(), chkOpenFolder.isSelected());

        progress.progressProperty().unbind();
        progress.progressProperty().bind(task.progressProperty());

        task.setOnSucceeded(e -> {
            progress.progressProperty().unbind();
            progress.setProgress(0);

            setUiRunning(false);
            status("변환 완료");
        });

        task.setOnFailed(e -> {
            progress.progressProperty().unbind();
            progress.setProgress(0);

            setUiRunning(false);
            Throwable ex = task.getException();
            status("오류 발생: " + (ex == null ? "(unknown)" : ex.getMessage()));
            if (ex != null) ex.printStackTrace();
        });

        Thread th = new Thread(task);
        th.setDaemon(true);
        th.start();
    }

    private void setUiRunning(boolean running) {
        // English comment: btnConvert.disableProperty() is bound. Control it via runningProperty only.
        runningProperty.set(running);

        btnPickHwp.setDisable(running);
        btnChangeHwp.setDisable(running);
        btnChangeOut.setDisable(running);
        chkOpenFolder.setDisable(running);
        chkAutoExpandLog.setDisable(running);

        // English comment: progress is bound to Task.progressProperty(). Do not call setProgress() here.
    }


    private boolean isHwpFile(File f) {
        if (f == null) return false;
        String name = f.getName().toLowerCase();
        return name.endsWith(".hwp");
    }

    private String makeOutName(String inputFileName) {
        String stem = inputFileName;
        int dot = stem.lastIndexOf('.');
        if (dot > 0) stem = stem.substring(0, dot);
        return stem + "_수정.xlsx";
    }

    private void log(String msg) {
        Platform.runLater(() -> {
            logArea.appendText(msg + "\n");
            logArea.positionCaret(logArea.getText().length());
        });
    }

    private void status(String msg) {
        Platform.runLater(() -> statusLabel.setText("상태: " + msg));
    }
}
