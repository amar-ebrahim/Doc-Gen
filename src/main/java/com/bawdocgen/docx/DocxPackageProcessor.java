package com.bawdocgen.docx;

import com.bawdocgen.api.DocumentGenerationException;
import com.bawdocgen.barcode.BarcodeImage;
import com.bawdocgen.barcode.BarcodeImageGenerator;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class DocxPackageProcessor {
    // Standard OOXML unit conversion: 914400 EMU per inch, 96 pixels per inch.
    private static final int EMU_PER_PIXEL = 9525;
    private static final Pattern RELATIONSHIP_ID = Pattern.compile("Id=\"rId(\\d+)\"");

    private final DocxTemplateSanitizer sanitizer = new DocxTemplateSanitizer();
    private final XmlPartTransformer xmlPartTransformer = new XmlPartTransformer();
    private final BarcodeImageGenerator barcodeImageGenerator = new BarcodeImageGenerator();

    public byte[] generate(byte[] templateBytes, Map<String, String> values,
                           Map<String, List<Map<String, String>>> tables) throws DocumentGenerationException {
        byte[] sanitizedTemplateBytes = sanitizer.sanitize(templateBytes);
        BarcodeRequest barcodeRequest = BarcodeRequest.from(values);

        LinkedHashMap<String, byte[]> entries = readZipEntries(sanitizedTemplateBytes);

        if (barcodeRequest != null) {
            insertBarcode(entries, barcodeRequest);
        }

        for (Map.Entry<String, byte[]> entry : entries.entrySet()) {
            if (isTransformableWordXmlPart(entry.getKey())) {
                entry.setValue(xmlPartTransformer.transform(entry.getValue(), values, tables));
            }
        }

        return writeZipEntries(entries);
    }

    public Set<String> extractPlaceholders(byte[] templateBytes) throws DocumentGenerationException {
        byte[] sanitizedTemplateBytes = sanitizer.sanitize(templateBytes);
        Set<String> placeholders = new LinkedHashSet<>();
        try (ZipInputStream inputStream = new ZipInputStream(new ByteArrayInputStream(sanitizedTemplateBytes))) {
            ZipEntry entry;
            while ((entry = inputStream.getNextEntry()) != null) {
                byte[] content = readAllBytes(inputStream);
                if (isTransformableWordXmlPart(entry.getName())) {
                    placeholders.addAll(xmlPartTransformer.extractPlaceholders(content));
                }
                inputStream.closeEntry();
            }
            return placeholders;
        } catch (IOException e) {
            throw new DocumentGenerationException("DOC-002", "Failed to inspect DOCX package", e);
        }
    }

    // Generates the barcode image, then embeds it into every transformable part that contains the
    // {{barcodeLocation}} placeholder: swaps the placeholder run for a <w:drawing>, adds an image
    // relationship scoped to that part's own .rels file, adds the media part, and registers the
    // png content type. Fails loudly if the placeholder isn't found anywhere, since a silently
    // ignored barcodeFlag would be a confusing surprise for a BAW author.
    private void insertBarcode(LinkedHashMap<String, byte[]> entries, BarcodeRequest barcodeRequest)
            throws DocumentGenerationException {
        BarcodeImage barcodeImage = barcodeImageGenerator.generate(
                barcodeRequest.getContent(), barcodeRequest.getType(), barcodeRequest.getWidth(), barcodeRequest.getHeight());

        String mediaFileName = nextMediaFileName(entries.keySet());
        int widthEmu = barcodeImage.getWidth() * EMU_PER_PIXEL;
        int heightEmu = barcodeImage.getHeight() * EMU_PER_PIXEL;
        boolean inserted = false;

        for (Map.Entry<String, byte[]> partEntry : new ArrayList<>(entries.entrySet())) {
            String partName = partEntry.getKey();
            if (!isTransformableWordXmlPart(partName)) {
                continue;
            }

            String relsPath = relsPathFor(partName);
            byte[] existingRels = entries.get(relsPath);
            String relationshipId = nextRelationshipId(existingRels);

            XmlPartTransformer.ImagePlaceholderResult result = xmlPartTransformer.insertImagePlaceholder(
                    partEntry.getValue(), barcodeRequest.getPlaceholderKey(), relationshipId,
                    widthEmu, heightEmu, mediaFileName);

            if (result.inserted) {
                entries.put(partName, result.xmlBytes);
                entries.put(relsPath, addImageRelationship(existingRels, relationshipId, mediaFileName));
                inserted = true;
            }
        }

        if (!inserted) {
            throw new DocumentGenerationException("DOC-005",
                    "barcodeFlag is true but placeholder {{" + barcodeRequest.getPlaceholderKey()
                            + "}} (barcodeLocation) was not found in the template");
        }

        entries.put("word/media/" + mediaFileName, barcodeImage.getPngBytes());
        entries.put("[Content_Types].xml", ensurePngContentType(entries.get("[Content_Types].xml")));
    }

    private String relsPathFor(String partName) {
        int slash = partName.lastIndexOf('/');
        String dir = partName.substring(0, slash);
        String fileName = partName.substring(slash + 1);
        return dir + "/_rels/" + fileName + ".rels";
    }

    private String nextMediaFileName(Set<String> existingEntryNames) {
        int index = 1;
        String candidate;
        do {
            candidate = "barcode" + index + ".png";
            index++;
        } while (existingEntryNames.contains("word/media/" + candidate));
        return candidate;
    }

    private String nextRelationshipId(byte[] existingRelsXml) {
        if (existingRelsXml == null) {
            return "rId1";
        }
        String xml = new String(existingRelsXml, StandardCharsets.UTF_8);
        Matcher matcher = RELATIONSHIP_ID.matcher(xml);
        int max = 0;
        while (matcher.find()) {
            max = Math.max(max, Integer.parseInt(matcher.group(1)));
        }
        return "rId" + (max + 1);
    }

    private byte[] addImageRelationship(byte[] existingRelsXml, String relationshipId, String mediaFileName)
            throws DocumentGenerationException {
        String relationship = "<Relationship Id=\"" + relationshipId + "\" "
                + "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" "
                + "Target=\"media/" + mediaFileName + "\"/>";

        if (existingRelsXml == null) {
            String newRels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + relationship + "</Relationships>";
            return newRels.getBytes(StandardCharsets.UTF_8);
        }

        String xml = new String(existingRelsXml, StandardCharsets.UTF_8);
        int closeIdx = xml.lastIndexOf("</Relationships>");
        if (closeIdx < 0) {
            throw new DocumentGenerationException("DOC-002", "Malformed relationships part in template");
        }
        String updated = xml.substring(0, closeIdx) + relationship + xml.substring(closeIdx);
        return updated.getBytes(StandardCharsets.UTF_8);
    }

    private byte[] ensurePngContentType(byte[] contentTypesXml) throws DocumentGenerationException {
        if (contentTypesXml == null) {
            throw new DocumentGenerationException("DOC-002", "Template is missing [Content_Types].xml");
        }
        String xml = new String(contentTypesXml, StandardCharsets.UTF_8);
        if (xml.contains("Extension=\"png\"")) {
            return contentTypesXml;
        }
        int closeIdx = xml.lastIndexOf("</Types>");
        if (closeIdx < 0) {
            throw new DocumentGenerationException("DOC-002", "Malformed [Content_Types].xml in template");
        }
        String updated = xml.substring(0, closeIdx)
                + "<Default Extension=\"png\" ContentType=\"image/png\"/>"
                + xml.substring(closeIdx);
        return updated.getBytes(StandardCharsets.UTF_8);
    }

    private boolean isTransformableWordXmlPart(String entryName) {
        if (entryName == null || !entryName.startsWith("word/") || !entryName.endsWith(".xml")) {
            return false;
        }
        return "word/document.xml".equals(entryName)
                || entryName.matches("word/header\\d+\\.xml")
                || entryName.matches("word/footer\\d+\\.xml")
                || "word/footnotes.xml".equals(entryName)
                || "word/endnotes.xml".equals(entryName)
                || "word/comments.xml".equals(entryName);
    }

    private LinkedHashMap<String, byte[]> readZipEntries(byte[] zipBytes) throws DocumentGenerationException {
        LinkedHashMap<String, byte[]> entries = new LinkedHashMap<>();
        try (ZipInputStream inputStream = new ZipInputStream(new ByteArrayInputStream(zipBytes))) {
            ZipEntry entry;
            while ((entry = inputStream.getNextEntry()) != null) {
                entries.put(entry.getName(), readAllBytes(inputStream));
                inputStream.closeEntry();
            }
            return entries;
        } catch (IOException e) {
            throw new DocumentGenerationException("DOC-002", "Failed to read DOCX package", e);
        }
    }

    private byte[] writeZipEntries(Map<String, byte[]> entries) throws DocumentGenerationException {
        try (ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();
             ZipOutputStream outputStream = new ZipOutputStream(byteOutputStream)) {
            for (Map.Entry<String, byte[]> entry : entries.entrySet()) {
                outputStream.putNextEntry(new ZipEntry(entry.getKey()));
                outputStream.write(entry.getValue());
                outputStream.closeEntry();
            }
            outputStream.finish();
            return byteOutputStream.toByteArray();
        } catch (IOException e) {
            throw new DocumentGenerationException("DOC-002", "Failed to generate DOCX package", e);
        }
    }

    private byte[] readAllBytes(ZipInputStream inputStream) throws IOException {
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
