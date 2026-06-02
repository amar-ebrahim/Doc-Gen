package com.bawdocgen.cli;

import com.bawdocgen.api.DocumentGenerator;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.charset.StandardCharsets;
import java.util.Base64;

public final class DocumentGeneratorCli {
    private DocumentGeneratorCli() {
    }

    public static void main(String[] args) throws Exception {
        if (args.length != 3) {
            System.err.println("Usage: DocumentGeneratorCli <template.docx|base64.txt> <data.json> <output.docx>");
            System.exit(2);
        }

        Path templatePath = Paths.get(args[0]);
        Path jsonPath = Paths.get(args[1]);
        Path outputPath = Paths.get(args[2]);

        byte[] template = resolveTemplate(Files.readAllBytes(templatePath));
        String json = resolveJson(new String(Files.readAllBytes(jsonPath), StandardCharsets.UTF_8));
        byte[] docx = DocumentGenerator.getInstance().generateDocx(template, json);

        Path parent = outputPath.toAbsolutePath().getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        Files.write(outputPath, docx);

        System.out.println("Generated DOCX:");
        System.out.println(outputPath.toAbsolutePath());
        System.out.println("Size: " + docx.length + " bytes");
    }

    // If the file doesn't start with PK (ZIP magic), treat it as base64-encoded DOCX.
    private static byte[] resolveTemplate(byte[] raw) {
        if (raw.length >= 2 && raw[0] == 'P' && raw[1] == 'K') {
            return raw;
        }
        String text = new String(raw, StandardCharsets.UTF_8).trim().replaceAll("\\s+", "");
        return Base64.getDecoder().decode(text);
    }

    // If the JSON has no "flat_mapping" key, wrap the entire object as the flat_mapping value.
    private static String resolveJson(String json) {
        String trimmed = json.trim();
        if (trimmed.contains("\"flat_mapping\"")) {
            return trimmed;
        }
        return "{\"flat_mapping\":" + trimmed + "}";
    }
}
