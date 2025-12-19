package org.example;

import org.example.HwpToJson;
import org.example.JsonToExcel;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;

public class AllInOneUI extends JFrame {

    private JTextField hwpField;
    private JTextField outputField;
    private JTextArea logArea;

    public AllInOneUI() {
        setTitle("조리지시서 만들기~~");
        setSize(650, 380);
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setLocationRelativeTo(null);

        initUI();
    }

    private void initUI() {
        JPanel panel = new JPanel(new BorderLayout(10, 10));
        panel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        JPanel top = new JPanel(new GridLayout(2, 1, 5, 5));
        top.add(createHwpPanel());
        top.add(createOutputPanel());

        panel.add(top, BorderLayout.NORTH);
        panel.add(createLogPanel(), BorderLayout.CENTER);
        panel.add(createConvertButton(), BorderLayout.SOUTH);

        add(panel);
    }

    private JPanel createHwpPanel() {
        JPanel p = new JPanel(new BorderLayout(5, 5));
        p.add(new JLabel("HWP 파일"), BorderLayout.WEST);

        hwpField = new JTextField();
        hwpField.setEditable(false);
        p.add(hwpField, BorderLayout.CENTER);

        JButton btn = new JButton("선택");
        btn.addActionListener(e -> chooseHwp());
        p.add(btn, BorderLayout.EAST);

        return p;
    }

    private JPanel createOutputPanel() {
        JPanel p = new JPanel(new BorderLayout(5, 5));
        p.add(new JLabel("출력 폴더"), BorderLayout.WEST);

        outputField = new JTextField(getDefaultOutputDir().toString());
        outputField.setEditable(false);
        p.add(outputField, BorderLayout.CENTER);

        JButton btn = new JButton("선택");
        btn.addActionListener(e -> chooseOutput());
        p.add(btn, BorderLayout.EAST);

        return p;
    }

    private JScrollPane createLogPanel() {
        logArea = new JTextArea();
        logArea.setEditable(false);
        return new JScrollPane(logArea);
    }

    private JButton createConvertButton() {
        JButton btn = new JButton("변환");
        btn.addActionListener(e -> convert());
        return btn;
    }

    private void chooseHwp() {
        JFileChooser fc = new JFileChooser();
        fc.setFileSelectionMode(JFileChooser.FILES_ONLY);

        if (fc.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            hwpField.setText(fc.getSelectedFile().getAbsolutePath());
        }
    }

    private void chooseOutput() {
        JFileChooser fc = new JFileChooser();
        fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        if (fc.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            outputField.setText(fc.getSelectedFile().getAbsolutePath());
        }
    }

    private void convert() {
        try {
            File hwp = new File(hwpField.getText());
            File outDir = new File(outputField.getText());

            if (!hwp.exists()) {
                log("HWP 파일이 없습니다.");
                return;
            }

            Path appHome = Paths.get(System.getProperty("user.dir")).toAbsolutePath();
            File inputDir = appHome.resolve("input").toFile();

            if (!inputDir.exists()) {
                log("input 폴더가 없습니다: " + inputDir.getAbsolutePath());
                return;
            }

            log("HWP 분석 시작...");
            String json;

            // 🔹 여기서 자동/수동 판별 가능
            boolean isExtended = hwp.getName().contains("시간연장");

            if (isExtended) {
                json = HwpToJson.convertToJsonString(hwp);
                JsonToExcel.convertFromJsonString(json, hwp.toPath(), outDir.toPath());
            } else {
                json = HwpToJsonGeneral.convertToJsonString(hwp);
                JsonToExcelGeneral.convertFromJsonString(json, hwp.toPath(), outDir.toPath());
            }

            log("변환 완료");

        } catch (Exception ex) {
            log("오류 발생: " + ex.getMessage());
            ex.printStackTrace();
        }
    }



    private Path getDefaultOutputDir() {
        Path base = Paths.get(System.getProperty("user.dir")).toAbsolutePath();
        Path out = base.resolve("output");
        out.toFile().mkdirs();
        return out;
    }

    private void log(String msg) {
        logArea.append(msg + "\n");
    }
}
