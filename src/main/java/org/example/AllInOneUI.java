package org.example;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Enumeration;

/**
 * Simple portable UI:
 * - Select HWP (any location)
 * - Select output XLSX file (any location)
 * - Uses input/ folder next to the program (appHome/input)
 */
public class AllInOneUI extends JFrame {

    private JTextField hwpField;
    private JTextField outputField;
    private JTextArea logArea;
    private boolean outputChosenByUser = false;

    public AllInOneUI() {
        setTitle("조리지시서 만들기~~");
        setSize(700, 420);
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setLocationRelativeTo(null);

        initUI();
    }

    private void initUI() {
        Font uiFont = new Font("맑은 고딕", Font.PLAIN, 13);
        setUIFont(uiFont);
        JPanel panel = new JPanel(new BorderLayout(10, 10));
        panel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        JPanel top = new JPanel(new GridLayout(2, 1, 6, 6));
        top.add(createHwpPanel());
        top.add(createOutputFilePanel());

        panel.add(top, BorderLayout.NORTH);
        panel.add(createLogPanel(), BorderLayout.CENTER);
        panel.add(createConvertButton(), BorderLayout.SOUTH);

        add(panel);

        // English comment: Show app home / input dir in log for troubleshooting.
        try {
            Path appHome = getAppHome();
            log("appHome: " + appHome);
            log("inputDir: " + appHome.resolve("input"));
        } catch (Exception e) {
            log("appHome resolve failed: " + e.getMessage());
        }
    }

    private JPanel createHwpPanel() {
        JPanel p = new JPanel(new BorderLayout(5, 5));
        p.add(new JLabel("HWP 파일"), BorderLayout.WEST);
        p.setPreferredSize(new Dimension(90, 28));

        hwpField = new JTextField();
        hwpField.setEditable(false);
        p.add(hwpField, BorderLayout.CENTER);

        JButton btn = new JButton("선택");
        btn.addActionListener(e -> chooseHwp());
        p.add(btn, BorderLayout.EAST);

        return p;
    }

    private JPanel createOutputFilePanel() {
        JPanel p = new JPanel(new BorderLayout(5, 5));
        p.add(new JLabel("출력 XLSX"), BorderLayout.WEST);
        p.setPreferredSize(new Dimension(90, 28));

        outputField = new JTextField();
        outputField.setEditable(false);
        p.add(outputField, BorderLayout.CENTER);

        JButton btn = new JButton("저장 위치 선택");
        btn.addActionListener(e -> chooseOutputXlsx());
        p.add(btn, BorderLayout.EAST);

        return p;
    }

    private JScrollPane createLogPanel() {
        logArea = new JTextArea();
        logArea.setEditable(false);
        logArea.setFont(new Font(Font.MONOSPACED, Font.PLAIN, 12));
        return new JScrollPane(logArea);
    }

    private JButton createConvertButton() {
        JButton btn = new JButton("변환");
        btn.setPreferredSize(new Dimension(120, 42));
        btn.setFont(new Font("맑은 고딕", Font.BOLD, 14));
        btn.addActionListener(e -> convert());
        return btn;
    }

    private void chooseHwp() {
        JFileChooser fc = new JFileChooser();
        fc.setFileSelectionMode(JFileChooser.FILES_ONLY);

        int r = showFileChooserLarge(fc, "HWP 파일 선택", false);
        if (r == JFileChooser.APPROVE_OPTION) {
            File f = fc.getSelectedFile();
            hwpField.setText(f.getAbsolutePath());

            // English comment: Auto-update output filename when HWP changes,
            // unless user explicitly selected output path.
            if (!outputChosenByUser) {
                File suggested = new File(f.getParentFile(), makeOutName(f.getName()));
                outputField.setText(suggested.getAbsolutePath());
            } else {
                // English comment: User chose output path; keep directory but update filename only.
                File current = new File(outputField.getText().trim());
                File parent = current.getParentFile();
                if (parent != null) {
                    File suggested = new File(parent, makeOutName(f.getName()));
                    outputField.setText(suggested.getAbsolutePath());
                }
            }
        }
    }

    /**
     * Output is a FILE chooser (not a directory chooser).
     * This avoids mandatory output/ folder creation and lets user save anywhere.
     */
    private void chooseOutputXlsx() {
        JFileChooser fc = new JFileChooser();
        fc.setFileSelectionMode(JFileChooser.FILES_ONLY);

        // English comment: Suggest default name based on selected HWP.
        String defaultName = "output.xlsx";
        if (hwpField.getText() != null && hwpField.getText().trim().length() > 0) {
            File hwp = new File(hwpField.getText().trim());
            defaultName = makeOutName(hwp.getName());
            if (hwp.getParentFile() != null) {
                fc.setCurrentDirectory(hwp.getParentFile());
            }
        }

        fc.setSelectedFile(new File(defaultName));

        int r = showFileChooserLarge(fc, "XLSX 저장 위치 선택", true);
        if (r == JFileChooser.APPROVE_OPTION) {
            File f = fc.getSelectedFile();

            // English comment: Ensure .xlsx extension.
            String path = f.getAbsolutePath();
            if (!path.toLowerCase().endsWith(".xlsx")) {
                f = new File(path + ".xlsx");
            }

            outputField.setText(f.getAbsolutePath());
            outputChosenByUser = true;
        }
    }

    private void convert() {
        try {
            // English comment: Validate input selection.
            if (hwpField.getText() == null || hwpField.getText().trim().length() == 0) {
                log("HWP 파일을 먼저 선택해주세요.");
                return;
            }
            if (outputField.getText() == null || outputField.getText().trim().length() == 0) {
                log("출력 XLSX 저장 위치를 선택해주세요.");
                return;
            }

            File hwp = new File(hwpField.getText().trim());
            File outXlsx = new File(outputField.getText().trim());

            if (!hwp.exists() || !hwp.isFile()) {
                log("HWP 파일이 없습니다: " + hwp.getAbsolutePath());
                return;
            }

            // English comment: Determine appHome (where the program is located) and resolve input folder next to it.
            Path appHome = getAppHome();
            File inputDir = appHome.resolve("input").toFile();

            log("resolved appHome: " + appHome);

            if (!inputDir.exists() || !inputDir.isDirectory()) {
                log("input 폴더가 없습니다: " + inputDir.getAbsolutePath());
                log("※ 프로그램 폴더 옆에 input/ 폴더를 두고 템플릿+allergy.png를 넣어주세요.");
                return;
            }

            // English comment: Many existing codes use "app.home" system property to resolve input resources.
            System.setProperty("app.home", appHome.toString());

            // English comment: Decide mode by filename (simple 1st rule only).
            boolean isExtended = hwp.getName().contains("시간연장");

            log("HWP 분석 시작...");
            log("INPUT  : " + hwp.getAbsolutePath());
            log("OUTPUT : " + outXlsx.getAbsolutePath());
            log("MODE   : " + (isExtended ? "EXTENDED" : "GENERAL"));

            String json;

            if (isExtended) {
                json = HwpToJson.convertToJsonString(hwp);
                JsonToExcel.convertFromJsonString(json, hwp.toPath(), outXlsx.toPath());
            } else {
                json = HwpToJsonGeneral.convertToJsonString(hwp);
                JsonToExcelGeneral.convertFromJsonString(json, hwp.toPath(), outXlsx.toPath());
            }

            log("변환 완료");

        } catch (Exception ex) {
            log("오류 발생: " + ex.getMessage());
            ex.printStackTrace();
        }
    }

    private int showFileChooserLarge(JFileChooser chooser, String title, boolean saveDialog) {
        JDialog dialog = new JDialog(this, title, true);
        dialog.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
        dialog.setLayout(new BorderLayout());
        dialog.add(chooser, BorderLayout.CENTER);

        dialog.setMinimumSize(new Dimension(980, 680));
        dialog.setSize(1080, 720);
        dialog.setLocationRelativeTo(this);
        dialog.setResizable(true);

        final int[] result = new int[]{JFileChooser.CANCEL_OPTION};

        chooser.addActionListener(e -> {
            String cmd = e.getActionCommand();
            if (JFileChooser.APPROVE_SELECTION.equals(cmd)) {
                result[0] = JFileChooser.APPROVE_OPTION;
                dialog.dispose();
            } else if (JFileChooser.CANCEL_SELECTION.equals(cmd)) {
                result[0] = JFileChooser.CANCEL_OPTION;
                dialog.dispose();
            }
        });

        dialog.setVisible(true);
        return result[0];
    }

    private Path getAppHome() throws Exception {
        // English comment: Resolve the location of the running jar within jpackage app-image.
        Path loc = Paths.get(AllInOneUI.class.getProtectionDomain().getCodeSource().getLocation().toURI())
                .toAbsolutePath();

        // English comment: If loc is a file (jar), use its parent folder.
        Path base = loc.toFile().isFile() ? loc.getParent() : loc;

        // English comment: 1) If input exists next to base, use base.
        Path input1 = base.resolve("input");
        if (input1.toFile().exists() && input1.toFile().isDirectory()) {
            return base;
        }

        // English comment: 2) jpackage app-image often places jars under <appHome>/app/.
        // If base ends with "app" and input exists at parent, use parent.
        Path parent = base.getParent();
        if (parent != null) {
            Path input2 = parent.resolve("input");
            if (input2.toFile().exists() && input2.toFile().isDirectory()) {
                return parent;
            }
        }

        // English comment: 3) Fallback to base (will fail with a clear log if input isn't found).
        return base;
    }

    private static void setUIFont(Font f) {
        Enumeration<Object> keys = UIManager.getDefaults().keys();
        while (keys.hasMoreElements()) {
            Object key = keys.nextElement();
            Object value = UIManager.get(key);
            if (value instanceof javax.swing.plaf.FontUIResource) {
                UIManager.put(key, new javax.swing.plaf.FontUIResource(f));
            }
        }
    }

    private String makeOutName(String inputFileName) {
        String stem = inputFileName;
        int dot = stem.lastIndexOf('.');
        if (dot > 0) {
            stem = stem.substring(0, dot);
        }
        return stem + "_수정.xlsx";
    }

    private void log(String msg) {
        logArea.append(msg + "\n");
        logArea.setCaretPosition(logArea.getDocument().getLength());
    }
}
