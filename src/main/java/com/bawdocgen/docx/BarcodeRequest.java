package com.bawdocgen.docx;

import com.bawdocgen.api.DocumentGenerationException;

import java.util.Map;

/**
 * Parsed view of the barcodeFlag/barcodeContent/barcodeLocation/barcodeType/barcodeWidth/barcodeHeight
 * flat_mapping keys used to opt a request into barcode generation.
 */
class BarcodeRequest {
    private final String content;
    private final String placeholderKey;
    private final String type;
    private final Integer width;
    private final Integer height;

    private BarcodeRequest(String content, String placeholderKey, String type, Integer width, Integer height) {
        this.content = content;
        this.placeholderKey = placeholderKey;
        this.type = type;
        this.width = width;
        this.height = height;
    }

    String getContent() {
        return content;
    }

    String getPlaceholderKey() {
        return placeholderKey;
    }

    String getType() {
        return type;
    }

    Integer getWidth() {
        return width;
    }

    Integer getHeight() {
        return height;
    }

    /**
     * Returns null when barcodeFlag is absent/falsy (barcode generation not requested).
     */
    static BarcodeRequest from(Map<String, String> values) throws DocumentGenerationException {
        if (!XmlPartTransformer.isTruthy(values.get("barcodeFlag"))) {
            return null;
        }

        String content = values.get("barcodeContent");
        if (content == null || content.trim().isEmpty()) {
            throw new DocumentGenerationException("DOC-005",
                    "barcodeFlag is true but barcodeContent is missing or empty");
        }

        String placeholderKey = values.get("barcodeLocation");
        if (placeholderKey == null || placeholderKey.trim().isEmpty()) {
            throw new DocumentGenerationException("DOC-005",
                    "barcodeFlag is true but barcodeLocation is missing or empty");
        }

        String type = values.get("barcodeType");
        Integer width = parsePositiveInt("barcodeWidth", values.get("barcodeWidth"));
        Integer height = parsePositiveInt("barcodeHeight", values.get("barcodeHeight"));
        return new BarcodeRequest(content, placeholderKey.trim(), type, width, height);
    }

    private static Integer parsePositiveInt(String field, String raw) throws DocumentGenerationException {
        if (raw == null || raw.trim().isEmpty()) {
            return null;
        }
        try {
            int value = Integer.parseInt(raw.trim());
            if (value <= 0) {
                throw new NumberFormatException();
            }
            return value;
        } catch (NumberFormatException e) {
            throw new DocumentGenerationException("DOC-005",
                    field + " must be a positive integer, got '" + raw + "'");
        }
    }
}
