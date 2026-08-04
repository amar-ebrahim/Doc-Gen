package com.bawdocgen.api;

import com.bawdocgen.json.FlatMappingJsonParser;

import java.security.CodeSource;
import java.util.Base64;
import java.util.Map;
import java.util.Set;

/**
 * BAW-friendly facade with Java Integration Service supported signatures.
 */
public class DocxGenerationService {
    private static final int JSON_PREVIEW_LENGTH = 300;
    private static final String BUILD_MARKER = "barcode-feature-2026-08-04-01";

    public DocxGenerationService() {
    }

    /**
     * Diagnostic entry point with no inputs, for confirming from BAW which jar build is
     * actually loaded at runtime (vs what was rebuilt locally). Call this, copy the returned
     * string back for review.
     */
    public String diagnostics() {
        StringBuilder sb = new StringBuilder();
        sb.append("BUILD_MARKER=").append(BUILD_MARKER).append('\n');
        sb.append("DocxGenerationService loaded from: ").append(codeSourceOf(DocxGenerationService.class)).append('\n');
        sb.append("DocumentGenerator loaded from: ").append(codeSourceOf(DocumentGenerator.class)).append('\n');

        sb.append("BarcodeImageGenerator (com.bawdocgen.barcode.BarcodeImageGenerator): ");
        try {
            Class<?> barcodeGenClass = Class.forName("com.bawdocgen.barcode.BarcodeImageGenerator");
            sb.append("FOUND, loaded from: ").append(codeSourceOf(barcodeGenClass)).append('\n');
        } catch (ClassNotFoundException e) {
            sb.append("NOT FOUND -- old jar without barcode support is loaded\n");
        }

        sb.append("Shaded ZXing (com.bawdocgen.shaded.zxing.BarcodeFormat): ");
        try {
            Class<?> zxingClass = Class.forName("com.bawdocgen.shaded.zxing.BarcodeFormat");
            sb.append("FOUND, loaded from: ").append(codeSourceOf(zxingClass)).append('\n');
        } catch (ClassNotFoundException e) {
            sb.append("NOT FOUND\n");
        }

        return sb.toString();
    }

    private String codeSourceOf(Class<?> clazz) {
        try {
            CodeSource codeSource = clazz.getProtectionDomain().getCodeSource();
            if (codeSource == null || codeSource.getLocation() == null) {
                return "unknown (no code source, likely bootstrap/system classloader)";
            }
            return codeSource.getLocation().toString();
        } catch (RuntimeException e) {
            return "error reading code source: " + e.getMessage();
        }
    }

    /**
     * Diagnostic twin of generateDocxBase64FromBytes with the same two parameters, for
     * confirming exactly what BAW is actually passing at runtime -- parsed key count, whether
     * barcode-related keys are present with what values, whether the template really contains
     * the requested placeholder, and the precise error if generation throws. Never throws itself;
     * all failures are captured in the returned string.
     */
    public String diagnoseGeneration(String templateBase64, String jsonPayload) {
        StringBuilder sb = new StringBuilder();
        sb.append("BUILD_MARKER=").append(BUILD_MARKER).append('\n');

        byte[] templateBytes;
        try {
            templateBytes = Base64.getDecoder().decode(templateBase64 == null ? "" : templateBase64);
            sb.append("templateBase64: decoded ").append(templateBytes.length).append(" bytes, ")
                    .append((templateBytes.length >= 2 && templateBytes[0] == 0x50 && templateBytes[1] == 0x4B)
                            ? "starts with PK (looks like a valid zip/docx)"
                            : "does NOT start with PK -- not a valid docx")
                    .append('\n');
        } catch (RuntimeException e) {
            sb.append("templateBase64: FAILED to base64-decode -- ").append(e.getMessage()).append('\n');
            return sb.toString();
        }

        try {
            Set<String> placeholders = DocumentGenerator.getInstance().extractPlaceholders(templateBytes);
            sb.append("Template placeholders found (").append(placeholders.size()).append("): ")
                    .append(placeholders).append('\n');
        } catch (DocumentGenerationException e) {
            sb.append("Failed to extract placeholders: code=").append(e.getErrorCode())
                    .append(" message=").append(e.getMessage()).append('\n');
        }

        Map<String, String> flatMapping;
        try {
            flatMapping = new FlatMappingJsonParser().parse(jsonPayload);
            sb.append("jsonPayload: parsed ").append(flatMapping.size()).append(" flat_mapping keys: ")
                    .append(flatMapping.keySet()).append('\n');
            sb.append("  barcodeFlag=").append(valueOrMissing(flatMapping, "barcodeFlag")).append('\n');
            sb.append("  barcodeContent=").append(valueOrMissing(flatMapping, "barcodeContent")).append('\n');
            sb.append("  barcodeLocation=").append(valueOrMissing(flatMapping, "barcodeLocation")).append('\n');
        } catch (DocumentGenerationException e) {
            sb.append("Failed to parse jsonPayload: code=").append(e.getErrorCode())
                    .append(" message=").append(e.getMessage()).append('\n');
            sb.append("jsonPayload preview: ").append(preview(jsonPayload)).append('\n');
            return sb.toString();
        }

        sb.append("Attempting actual generateDocxBase64FromBytes call...\n");
        try {
            String result = generateDocxBase64FromBytes(templateBase64, jsonPayload);
            byte[] resultBytes = Base64.getDecoder().decode(result);
            sb.append("SUCCESS: generated ").append(resultBytes.length).append(" bytes.\n");
        } catch (DocxGenerationServiceException e) {
            sb.append("THREW: ").append(e.getMessage()).append('\n');
        } catch (RuntimeException e) {
            sb.append("THREW unexpected: ").append(e.getClass().getName()).append(": ").append(e.getMessage()).append('\n');
        }

        return sb.toString();
    }

    private String valueOrMissing(Map<String, String> map, String key) {
        return map.containsKey(key) ? "'" + map.get(key) + "'" : "MISSING";
    }

    public String generateDocxBase64FromBytes(String templateBase64, String jsonPayload) {
        try {
            byte[] templateBytes = Base64.getDecoder().decode(templateBase64);
            byte[] docx = DocumentGenerator.getInstance().generateDocx(templateBytes, jsonPayload);
            return Base64.getEncoder().encodeToString(docx);
        } catch (DocumentGenerationException e) {
            throw new DocxGenerationServiceException(formatError("<direct-bytes>", jsonPayload, e), e);
        } catch (RuntimeException e) {
            throw new DocxGenerationServiceException(formatUnexpectedError("<direct-bytes>", jsonPayload, e), e);
        }
    }

    public String generateDocxBase64(String templateName, String jsonPayload) {
        try {
            byte[] docx = DocumentGenerator.getInstance().generateDocx(templateName, jsonPayload);
            return Base64.getEncoder().encodeToString(docx);
        } catch (DocumentGenerationException e) {
            throw new DocxGenerationServiceException(formatError(templateName, jsonPayload, e), e);
        } catch (RuntimeException e) {
            throw new DocxGenerationServiceException(formatUnexpectedError(templateName, jsonPayload, e), e);
        }
    }

    private String formatError(String templateName, String jsonPayload, DocumentGenerationException e) {
        return "BAW DOCX generation failed"
                + " | code=" + valueOrEmpty(e.getErrorCode())
                + " | message=" + valueOrEmpty(e.getMessage())
                + " | templateName=" + valueOrEmpty(templateName)
                + " | jsonPreview=" + preview(jsonPayload)
                + " | rootCause=" + rootCauseSummary(e);
    }

    private String formatUnexpectedError(String templateName, String jsonPayload, RuntimeException e) {
        return "BAW DOCX generation failed"
                + " | code=DOC-999"
                + " | message=Unexpected runtime error"
                + " | templateName=" + valueOrEmpty(templateName)
                + " | jsonPreview=" + preview(jsonPayload)
                + " | rootCause=" + rootCauseSummary(e);
    }

    private String rootCauseSummary(Throwable throwable) {
        Throwable root = throwable;
        while (root.getCause() != null) {
            root = root.getCause();
        }
        return root.getClass().getSimpleName() + ": " + valueOrEmpty(root.getMessage());
    }

    private String preview(String value) {
        String normalized = valueOrEmpty(value).replace('\n', ' ').replace('\r', ' ').trim();
        if (normalized.length() <= JSON_PREVIEW_LENGTH) {
            return normalized;
        }
        return normalized.substring(0, JSON_PREVIEW_LENGTH) + "...";
    }

    private String valueOrEmpty(String value) {
        return value == null ? "" : value;
    }

    public static class DocxGenerationServiceException extends RuntimeException {
        public DocxGenerationServiceException(String message, Throwable cause) {
            super(message, cause);
        }
    }
}
