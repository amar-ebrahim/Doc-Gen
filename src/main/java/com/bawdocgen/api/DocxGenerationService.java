package com.bawdocgen.api;

import java.util.Base64;

/**
 * BAW-friendly facade with Java Integration Service supported signatures.
 */
public class DocxGenerationService {
    private static final int JSON_PREVIEW_LENGTH = 300;

    public DocxGenerationService() {
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
