package org.example;

import javax.swing.SwingUtilities;

public class Main {
    public static void main(String[] args) {
        // English comment: Ensure app.home is available for template resolution
        System.setProperty("app.home", java.nio.file.Paths.get(System.getProperty("user.dir")).toAbsolutePath().toString());

        SwingUtilities.invokeLater(() -> new AllInOneUI().setVisible(true));
    }
}
