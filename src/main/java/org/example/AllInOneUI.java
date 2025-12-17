package org.example;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.awt.datatransfer.DataFlavor;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

public class AllInOneUI extends JFrame {

    private final JTextField txtInput = new JTextField();
    private final JTextField txtOutputDir = new JTextField();
    private final JTextArea log = new JTextArea(12, 60);

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            AllInOneUI ui = new AllInOneUI();
            ui.setVisible(true);

            // If launched with a file path (drag and drop or file association)
            if (args != null && args.length >= 1) {
                String p = args[0];
                if (p != null && p.trim().length() > 0) {
                    ui.txtInput.setText(p.trim());
                }
            }
        });
    }

    public AllInOneUI() {
        setTitle("Menu HWP \u2192 Excel Converter");
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        setSize(820, 420);
        setLocationRelativeTo(null);

        JPanel root = new JPanel(new BorderLayout(10, 10));
        root.setBorder(new EmptyBorder(12, 12, 12, 12));
        setContentPane(root);

        JPanel top = new JPanel(new GridBagLayout());
        root.add(top, BorderLayout.NORTH);

        GridBagConstraints gc = new GridBagConstraints();
        gc.insets = new Insets(6, 6, 6, 6);
        gc.fill = GridBagConstraints.HORIZONTAL;

        // Input row
        gc.gridx = 0; gc.gridy = 0; gc.weightx = 0;
        top.add(new JLabel("Input HWP"), gc);

        gc.gridx = 1; gc.gridy = 0; gc.weightx = 1;
        top.add(txtInput, gc);

        JButton btnPick = new JButton("파일 선택...");
        gc.gridx = 2; gc.gridy = 0; gc.weightx = 0;
        top.add(btnPick, gc);

        // Output row
        gc.gridx = 0; gc.gridy = 1; gc.weightx = 0;
        top.add(new JLabel("Output 경로"), gc);

        // Default output to Desktop
        Path desktop = Paths.get(System.getProperty("user.home"), "Desktop");
        txtOutputDir.setText(desktop.toAbsolutePath().toString());

        gc.gridx = 1; gc.gridy = 1; gc.weightx = 1;
        top.add(txtOutputDir, gc);

        JButton btnPickOut = new JButton("폴더 선택...");
        gc.gridx = 2; gc.gridy = 1; gc.weightx = 0;
        top.add(btnPickOut, gc);

        // Buttons row
        JPanel btns = new JPanel(new FlowLayout(FlowLayout.RIGHT, 10, 0));
        JButton btnRun = new JButton("변환 시작");
        btns.add(btnRun);

        gc.gridx = 0; gc.gridy = 2; gc.gridwidth = 3; gc.weightx = 1;
        top.add(btns, gc);

        // Log area
        log.setEditable(false);
        JScrollPane sp = new JScrollPane(log);
        root.add(sp, BorderLayout.CENTER);

        // Drag and drop support
        root.setTransferHandler(new TransferHandler() {
            @Override
            public boolean canImport(TransferSupport support) {
                return support.isDataFlavorSupported(DataFlavor.javaFileListFlavor);
            }

            @Override
            public boolean importData(TransferSupport support) {
                try {
                    Object data = support.getTransferable().getTransferData(DataFlavor.javaFileListFlavor);
                    if (data instanceof List) {
                        List list = (List) data;
                        if (!list.isEmpty()) {
                            Object first = list.get(0);
                            if (first instanceof File) {
                                File f = (File) first;
                                txtInput.setText(f.getAbsolutePath());
                                appendLog("Dropped: " + f.getAbsolutePath());
                                return true;
                            }
                        }
                    }
                } catch (Exception e) {
                    appendLog("Drag and drop error: " + e.getMessage());
                }
                return false;
            }
        });

        // Handlers
        btnPick.addActionListener(e -> pickInputFile());
        btnPickOut.addActionListener(e -> pickOutputDir());
        btnRun.addActionListener(e -> runConvert());
    }

    private void pickInputFile() {
        JFileChooser fc = new JFileChooser();
        fc.setDialogTitle("HWP 파일 선택");
        int r = fc.showOpenDialog(this);
        if (r == JFileChooser.APPROVE_OPTION) {
            File f = fc.getSelectedFile();
            if (f != null) txtInput.setText(f.getAbsolutePath());
        }
    }

    private void pickOutputDir() {
        JFileChooser fc = new JFileChooser();
        fc.setDialogTitle("Output 폴더 선택");
        fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int r = fc.showOpenDialog(this);
        if (r == JFileChooser.APPROVE_OPTION) {
            File f = fc.getSelectedFile();
            if (f != null) txtOutputDir.setText(f.getAbsolutePath());
        }
    }

    private void runConvert() {
        String inStr = txtInput.getText();
        if (inStr == null || inStr.trim().length() == 0) {
            JOptionPane.showMessageDialog(this, "Input HWP를 선택하세요.");
            return;
        }

        Path in = Paths.get(inStr.trim()).toAbsolutePath();
        if (!Files.exists(in)) {
            JOptionPane.showMessageDialog(this, "파일이 없습니다:\n" + in);
            return;
        }

        Path outXlsx;
        try {
            outXlsx = buildOutputXlsxPath(in);
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Output 경로 처리 실패:\n" + ex.getMessage());
            return;
        }

        appendLog("INPUT : " + in.toAbsolutePath());
        appendLog("OUTPUT: " + outXlsx.toAbsolutePath());

        // Run in background thread to avoid UI freeze
        new Thread(() -> {
            try {
                // Call AllInOne with 2 args: inputHwp, outputXlsx
                String[] args = new String[]{in.toString(), outXlsx.toString()};
                AllInOne.main(args);

                appendLog("DONE: " + outXlsx.toAbsolutePath());
                SwingUtilities.invokeLater(() ->
                        JOptionPane.showMessageDialog(this, "완료!\n" + outXlsx.toAbsolutePath())
                );
            } catch (Exception ex) {
                appendLog("ERROR: " + ex.getMessage());
                SwingUtilities.invokeLater(() ->
                        JOptionPane.showMessageDialog(this, "에러:\n" + ex.getMessage())
                );
            }
        }).start();
    }

    private Path buildOutputXlsxPath(Path inputHwp) throws Exception {
        String outStr = txtOutputDir.getText();
        if (outStr == null) outStr = "";
        outStr = outStr.trim();

        // Empty means Desktop default
        if (outStr.length() == 0) {
            Path desktop = Paths.get(System.getProperty("user.home"), "Desktop").toAbsolutePath();
            Files.createDirectories(desktop);
            return desktop.resolve(makeOutName(inputHwp.getFileName().toString()));
        }

        // If user ends with a separator, treat as folder
        if (outStr.endsWith("\\") || outStr.endsWith("/")) {
            Path outDir = Paths.get(outStr).toAbsolutePath();
            Files.createDirectories(outDir);
            return outDir.resolve(makeOutName(inputHwp.getFileName().toString()));
        }

        Path outPath = Paths.get(outStr).toAbsolutePath();

        // Existing directory -> file inside it
        if (Files.exists(outPath) && Files.isDirectory(outPath)) {
            Files.createDirectories(outPath);
            return outPath.resolve(makeOutName(inputHwp.getFileName().toString()));
        }

        // Otherwise treat as file path
        String lower = outPath.toString().toLowerCase();
        if (!lower.endsWith(".xlsx")) {
            outPath = Paths.get(outPath.toString() + ".xlsx").toAbsolutePath();
        }

        Path parent = outPath.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }

        return outPath;
    }

    private String makeOutName(String inputFileName) {
        String name = inputFileName;
        int dot = name.lastIndexOf('.');
        if (dot > 0) name = name.substring(0, dot);
        return name + "_수정.xlsx";
    }

    private void appendLog(String s) {
        SwingUtilities.invokeLater(() -> {
            log.append(s + "\n");
            log.setCaretPosition(log.getDocument().getLength());
        });
    }
}
