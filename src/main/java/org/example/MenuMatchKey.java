package org.example;

final class MenuMatchKey {
    private MenuMatchKey() {}

    /** Keep allergen numbers; only ignore layout whitespace and template step metadata. */
    static String withAllergens(String name) {
        return MenuNamePrefix.withoutStep(name).replaceAll("[\\s\\p{Z}\\u200B]+", "");
    }
}
