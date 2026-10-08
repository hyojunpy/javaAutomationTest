package org.example;

import java.util.regex.Pattern;

/** Template step labels are metadata, not part of the recipe lookup key. */
final class MenuNamePrefix {
    private static final Pattern STEP = Pattern.compile("^\\s*step\\s*[1-3](?!\\d)\\s*",
            Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CHARACTER_CLASS);
    private MenuNamePrefix() {}

    static String withoutStep(String name) {
        return name == null ? "" : STEP.matcher(name).replaceFirst("");
    }
}
