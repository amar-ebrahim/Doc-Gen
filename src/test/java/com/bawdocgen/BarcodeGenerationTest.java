package com.bawdocgen;

import com.bawdocgen.api.DocumentGenerationException;
import com.bawdocgen.docx.DocxPackageProcessor;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class BarcodeGenerationTest {
    @Test
    void embedsBarcodeImageAtPlaceholderLocation() throws Exception {
        byte[] docx = minimalDocx("<w:p><w:r><w:t>{{BarcodeHere}}</w:t></w:r></w:p>");
        Map<String, String> values = new LinkedHashMap<>();
        values.put("barcodeFlag", "true");
        values.put("barcodeContent", "123456789");
        values.put("barcodeLocation", "BarcodeHere");

        byte[] generated = new DocxPackageProcessor().generate(docx, values, Collections.emptyMap());

        String documentXml = docxEntryAsString(generated, "word/document.xml");
        assertFalse(documentXml.contains("{{BarcodeHere}}"));
        assertTrue(documentXml.contains("<w:drawing"));
        assertTrue(documentXml.contains("r:embed=\"rId1\""));

        String rels = docxEntryAsString(generated, "word/_rels/document.xml.rels");
        assertTrue(rels.contains("Id=\"rId1\""));
        assertTrue(rels.contains("Target=\"media/barcode1.png\""));
        assertTrue(rels.contains("relationships/image"));

        byte[] png = docxEntryAsBytes(generated, "word/media/barcode1.png");
        assertTrue(png.length > 0);

        String contentTypes = docxEntryAsString(generated, "[Content_Types].xml");
        assertTrue(contentTypes.contains("Extension=\"png\""));
    }

    @Test
    void supportsQrCodeType() throws Exception {
        byte[] docx = minimalDocx("<w:p><w:r><w:t>{{Qr}}</w:t></w:r></w:p>");
        Map<String, String> values = new LinkedHashMap<>();
        values.put("barcodeFlag", "true");
        values.put("barcodeContent", "https://example.com/loan/12345");
        values.put("barcodeLocation", "Qr");
        values.put("barcodeType", "QR");

        byte[] generated = new DocxPackageProcessor().generate(docx, values, Collections.emptyMap());
        String documentXml = docxEntryAsString(generated, "word/document.xml");
        assertTrue(documentXml.contains("<w:drawing"));
    }

    @Test
    void doesNothingWhenBarcodeFlagIsAbsent() throws Exception {
        byte[] docx = minimalDocx("<w:p><w:r><w:t>{{Name}}</w:t></w:r></w:p>");
        Map<String, String> values = Collections.singletonMap("Name", "Maya");

        byte[] generated = new DocxPackageProcessor().generate(docx, values, Collections.emptyMap());

        assertTrue(docxEntryAsString(generated, "word/document.xml").contains("Maya"));
        assertNull(findEntry(generated, "word/media/barcode1.png"));
    }

    @Test
    void throwsWhenPlaceholderLocationIsMissingFromTemplate() throws Exception {
        byte[] docx = minimalDocx("<w:p><w:r><w:t>No barcode placeholder here</w:t></w:r></w:p>");
        Map<String, String> values = new LinkedHashMap<>();
        values.put("barcodeFlag", "true");
        values.put("barcodeContent", "123456789");
        values.put("barcodeLocation", "DoesNotExist");

        DocumentGenerationException exception = assertThrows(DocumentGenerationException.class,
                () -> new DocxPackageProcessor().generate(docx, values, Collections.emptyMap()));
        assertEquals("DOC-005", exception.getErrorCode());
    }

    @Test
    void throwsWhenBarcodeContentIsMissing() {
        Map<String, String> values = new LinkedHashMap<>();
        values.put("barcodeFlag", "true");
        values.put("barcodeLocation", "BarcodeHere");

        DocumentGenerationException exception = assertThrows(DocumentGenerationException.class,
                () -> new DocxPackageProcessor().generate(
                        minimalDocx("<w:p><w:r><w:t>{{BarcodeHere}}</w:t></w:r></w:p>"), values, Collections.emptyMap()));
        assertEquals("DOC-005", exception.getErrorCode());
    }

    private byte[] minimalDocx(String body) throws Exception {
        String documentXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
                + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                + "<w:body>" + body + "</w:body></w:document>";
        String contentTypes = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\""
                + " ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "</Types>";
        try (ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();
             ZipOutputStream zipOutputStream = new ZipOutputStream(byteOutputStream)) {
            writeEntry(zipOutputStream, "[Content_Types].xml", contentTypes);
            writeEntry(zipOutputStream, "word/document.xml", documentXml);
            zipOutputStream.finish();
            return byteOutputStream.toByteArray();
        }
    }

    private void writeEntry(ZipOutputStream zipOutputStream, String name, String content) throws Exception {
        zipOutputStream.putNextEntry(new ZipEntry(name));
        zipOutputStream.write(content.getBytes(StandardCharsets.UTF_8));
        zipOutputStream.closeEntry();
    }

    private String docxEntryAsString(byte[] docx, String entryName) throws Exception {
        return new String(docxEntryAsBytes(docx, entryName), StandardCharsets.UTF_8);
    }

    private byte[] docxEntryAsBytes(byte[] docx, String entryName) throws Exception {
        byte[] found = findEntry(docx, entryName);
        if (found == null) {
            throw new IllegalArgumentException("Missing DOCX entry " + entryName);
        }
        return found;
    }

    private byte[] findEntry(byte[] docx, String entryName) throws Exception {
        try (ZipInputStream inputStream = new ZipInputStream(new java.io.ByteArrayInputStream(docx))) {
            ZipEntry entry;
            while ((entry = inputStream.getNextEntry()) != null) {
                if (entryName.equals(entry.getName())) {
                    byte[] buffer = new byte[8192];
                    int read;
                    try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
                        while ((read = inputStream.read(buffer)) != -1) {
                            outputStream.write(buffer, 0, read);
                        }
                        return outputStream.toByteArray();
                    }
                }
                inputStream.closeEntry();
            }
        }
        return null;
    }
}
