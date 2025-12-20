package org.example;

import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * Resolve app home directory similar to Swing version.
 */
public class AppHomeResolver {

    public static Path getAppHome(Class<?> anchorClass) throws Exception {
        // English comment: Resolve the location of the running jar within jpackage app-image.
        Path loc = Paths.get(anchorClass.getProtectionDomain().getCodeSource().getLocation().toURI()).toAbsolutePath();

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

        // English comment: 3) Fallback to base.
        return base;
    }
}
