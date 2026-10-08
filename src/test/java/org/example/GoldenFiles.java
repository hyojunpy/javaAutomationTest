package org.example;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.w3c.dom.*;
import javax.xml.parsers.DocumentBuilderFactory;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.security.MessageDigest;
import java.util.*;
import java.util.zip.ZipFile;

/** Hash canonical XML and binary assets, ignoring ZIP timestamps and document properties. */
final class GoldenFiles {
    static final ObjectMapper JSON = new ObjectMapper();

    static Map<String, String> workbook(Path xlsx) throws Exception {
        Map<String, String> parts = new TreeMap<>();
        try (ZipFile zip = new ZipFile(xlsx.toFile())) {
            for (var entry : Collections.list(zip.entries())) {
                if (entry.isDirectory() || !entry.getName().startsWith("xl/")) continue;
                byte[] data;
                try (var stream = zip.getInputStream(entry)) { data = stream.readAllBytes(); }
                if (entry.getName().endsWith(".xml") || entry.getName().endsWith(".rels")) {
                    DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
                    factory.setNamespaceAware(true);
                    factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
                    var document = factory.newDocumentBuilder().parse(new java.io.ByteArrayInputStream(data));
                    StringBuilder normalized = new StringBuilder();
                    canonical(document.getDocumentElement(), normalized);
                    data = normalized.toString().getBytes(StandardCharsets.UTF_8);
                }
                parts.put(entry.getName(), HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(data)));
            }
        }
        return parts;
    }

    private static String name(Node node) {
        return "{" + Objects.toString(node.getNamespaceURI(), "") + "}" + node.getLocalName();
    }

    private static void token(String value, StringBuilder out) { out.append(value.length()).append(':').append(value); }

    private static void canonical(Node node, StringBuilder out) {
        if (node.getNodeType() == Node.ELEMENT_NODE) {
            out.append('E'); token(name(node), out);
            Map<String, String> attributes = new TreeMap<>();
            for (int i = 0; i < node.getAttributes().getLength(); i++) {
                Node attr = node.getAttributes().item(i);
                if (!"http://www.w3.org/2000/xmlns/".equals(attr.getNamespaceURI())) attributes.put(name(attr), attr.getNodeValue());
            }
            attributes.forEach((key, value) -> { out.append('A'); token(key, out); token(value, out); });
            for (Node child = node.getFirstChild(); child != null; child = child.getNextSibling()) canonical(child, out);
            out.append('Z');
        } else if (node.getNodeType() == Node.TEXT_NODE || node.getNodeType() == Node.CDATA_SECTION_NODE) {
            String value = node.getNodeValue();
            if (!value.isBlank() || "t".equals(node.getParentNode().getLocalName())) { out.append('T'); token(value, out); }
        }
    }

    static JsonNode capture(Path source, Path workbook) throws Exception {
        boolean extended = source.getFileName().toString().contains("시간연장");
        String parsed = extended ? HwpToJson.convertToJsonString(source) : HwpToJsonGeneral.convertToJsonString(source);
        var result = JSON.createObjectNode();
        result.set("plan", JSON.readTree(parsed));
        result.set("workbookParts", JSON.valueToTree(workbook(workbook)));
        return result;
    }

    /** Maintenance tool: use only a reviewed, original JAR and its outputs to capture baselines. */
    public static void main(String[] args) throws Exception {
        Path input = Path.of(args[0]), outputs = Path.of(args[1]), fixtures = Path.of(args[2]);
        Files.createDirectories(fixtures);
        try (var files = Files.list(input)) {
            for (Path source : files.filter(p -> p.toString().endsWith(".hwp")).sorted().toList()) {
                Path output = outputs.resolve(source.getFileName() + ".xlsx");
                JSON.writerWithDefaultPrettyPrinter().writeValue(fixtures.resolve(source.getFileName()+".golden.json").toFile(), capture(source, output));
                System.out.println("Captured " + source.getFileName());
            }
        }
    }
}
