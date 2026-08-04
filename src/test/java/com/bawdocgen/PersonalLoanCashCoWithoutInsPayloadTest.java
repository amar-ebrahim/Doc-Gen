package com.bawdocgen;

import com.bawdocgen.api.DocxGenerationService;
import com.bawdocgen.api.DocumentGenerator;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Base64;
import java.util.LinkedHashSet;
import java.util.Set;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class PersonalLoanCashCoWithoutInsPayloadTest {
    private static final String TEMPLATE_NAME = "Personal Loan Cash-Co without INS";
    private static final String PAYLOAD_PATH = "Personal_Loan_Cash-Co_without_INS_test_payload.json";

    @Test
    void payloadCoversTemplateAndChecksBothGuaranteeBoxes() throws Exception {
        byte[] template = resourceBytes("/templates/" + TEMPLATE_NAME + ".docx");
        String json = new String(Files.readAllBytes(Paths.get(PAYLOAD_PATH)), StandardCharsets.UTF_8);

        Set<String> placeholders = DocumentGenerator.getInstance().extractPlaceholders(template);
        placeholders.add("cashcobox");
        placeholders.add("PGBox");
        assertEquals(placeholders, payloadKeys(json));

        String docxBase64 = new DocxGenerationService().generateDocxBase64(TEMPLATE_NAME, json);
        byte[] docx = Base64.getDecoder().decode(docxBase64);
        String documentXml = docxEntry(docx, "word/document.xml");

        assertTrue(docx.length > 1000);
        assertTrue(new String(docx, 0, 2, StandardCharsets.US_ASCII).startsWith("PK"));
        assertFalse(documentXml.contains("{{"));
        assertCheckboxChecked(documentXml, "cashcobox");
        assertCheckboxChecked(documentXml, "PGBox");
        assertCheckboxContainsOnlySymbol(documentXml, "cashcobox");
        assertCheckboxContainsOnlySymbol(documentXml, "PGBox");
        assertTextOutsideCheckbox(documentXml, "cashcobox", "001-987654-002");
        assertTextOutsideCheckbox(documentXml, "PGBox", "Nadim Karim");
    }

    private Set<String> payloadKeys(String json) throws Exception {
        JsonNode root = new ObjectMapper().readTree(json);
        Set<String> keys = new LinkedHashSet<>();
        root.get("flat_mapping").fieldNames().forEachRemaining(keys::add);
        root.get("tables").fields().forEachRemaining(table -> {
            JsonNode rows = table.getValue();
            if (rows.isArray() && rows.size() > 0) {
                rows.get(0).fieldNames()
                        .forEachRemaining(field -> keys.add(table.getKey() + "[]." + field));
            }
        });
        return keys;
    }

    private void assertCheckboxChecked(String documentXml, String key) {
        String checkboxXml = checkboxXml(documentXml, key);

        assertTrue(checkboxXml.matches("(?s).*<w14:checked[^>]*w14:val=\"1\"[^>]*/>.*"),
                key + " should be checked");
        assertTrue(checkboxXml.contains("☒"), key + " should display the checked symbol");
    }

    private void assertCheckboxContainsOnlySymbol(String documentXml, String key) {
        String checkboxXml = checkboxXml(documentXml, key);
        String visibleText = checkboxXml.replaceAll("(?s)<[^>]+>", "").trim();
        assertEquals("☒", visibleText, key + " content control must contain only its symbol");
    }

    private void assertTextOutsideCheckbox(String documentXml, String key, String expectedText) {
        String checkboxXml = checkboxXml(documentXml, key);
        assertFalse(checkboxXml.contains(expectedText), expectedText + " must remain outside " + key);
        assertTrue(documentXml.contains(expectedText), "Generated document should contain " + expectedText);
    }

    private String checkboxXml(String documentXml, String key) {
        int tagIndex = documentXml.indexOf("w:val=\"" + key + "\"");
        assertTrue(tagIndex >= 0, "Missing checkbox tag for " + key);

        int checkboxStart = documentXml.lastIndexOf("<w:sdt", tagIndex);
        int checkboxEnd = documentXml.indexOf("</w:sdt>", tagIndex) + "</w:sdt>".length();
        return documentXml.substring(checkboxStart, checkboxEnd);
    }

    private byte[] resourceBytes(String path) throws Exception {
        try (InputStream inputStream = getClass().getResourceAsStream(path)) {
            if (inputStream == null) {
                throw new IllegalArgumentException("Missing test resource " + path);
            }
            byte[] buffer = new byte[8192];
            int read;
            try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
                while ((read = inputStream.read(buffer)) != -1) {
                    outputStream.write(buffer, 0, read);
                }
                return outputStream.toByteArray();
            }
        }
    }

    private String docxEntry(byte[] docx, String entryName) throws Exception {
        try (ZipInputStream inputStream = new ZipInputStream(new ByteArrayInputStream(docx))) {
            ZipEntry entry;
            while ((entry = inputStream.getNextEntry()) != null) {
                if (entryName.equals(entry.getName())) {
                    byte[] buffer = new byte[8192];
                    int read;
                    try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
                        while ((read = inputStream.read(buffer)) != -1) {
                            outputStream.write(buffer, 0, read);
                        }
                        return new String(outputStream.toByteArray(), StandardCharsets.UTF_8);
                    }
                }
                inputStream.closeEntry();
            }
        }
        throw new IllegalArgumentException("Missing DOCX entry " + entryName);
    }
}
