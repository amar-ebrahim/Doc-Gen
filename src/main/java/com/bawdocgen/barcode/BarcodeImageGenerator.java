package com.bawdocgen.barcode;

import com.bawdocgen.api.DocumentGenerationException;
import com.google.zxing.BarcodeFormat;
import com.google.zxing.EncodeHintType;
import com.google.zxing.MultiFormatWriter;
import com.google.zxing.WriterException;
import com.google.zxing.common.BitMatrix;
import com.google.zxing.qrcode.decoder.ErrorCorrectionLevel;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.EnumMap;
import java.util.Locale;
import java.util.Map;

public class BarcodeImageGenerator {
    private static final int DEFAULT_CODE128_WIDTH = 300;
    private static final int DEFAULT_CODE128_HEIGHT = 50;
    private static final int DEFAULT_QR_SIZE = 200;

    public BarcodeImage generate(String content, String barcodeType, Integer requestedWidth, Integer requestedHeight)
            throws DocumentGenerationException {
        BarcodeFormat format = resolveFormat(barcodeType);
        int width = requestedWidth != null ? requestedWidth : defaultWidth(format);
        int height = requestedHeight != null ? requestedHeight : defaultHeight(format);

        Map<EncodeHintType, Object> hints = new EnumMap<>(EncodeHintType.class);
        hints.put(EncodeHintType.MARGIN, 1);
        if (format == BarcodeFormat.QR_CODE) {
            hints.put(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.M);
        }

        try {
            BitMatrix matrix = new MultiFormatWriter().encode(content, format, width, height, hints);
            BufferedImage image = toBufferedImage(matrix);
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ImageIO.write(image, "png", out);
            return new BarcodeImage(out.toByteArray(), matrix.getWidth(), matrix.getHeight());
        } catch (WriterException | IOException e) {
            throw new DocumentGenerationException("DOC-005", "Failed to generate " + format + " barcode image", e);
        }
    }

    private BarcodeFormat resolveFormat(String barcodeType) throws DocumentGenerationException {
        if (barcodeType == null || barcodeType.trim().isEmpty()) {
            return BarcodeFormat.CODE_128;
        }
        String normalized = barcodeType.trim().toUpperCase(Locale.ROOT).replace("-", "_").replace(" ", "_");
        switch (normalized) {
            case "CODE128":
            case "CODE_128":
                return BarcodeFormat.CODE_128;
            case "QR":
            case "QRCODE":
            case "QR_CODE":
                return BarcodeFormat.QR_CODE;
            default:
                throw new DocumentGenerationException("DOC-005",
                        "Unsupported barcodeType '" + barcodeType + "'; expected CODE128 or QR");
        }
    }

    private int defaultWidth(BarcodeFormat format) {
        return format == BarcodeFormat.QR_CODE ? DEFAULT_QR_SIZE : DEFAULT_CODE128_WIDTH;
    }

    private int defaultHeight(BarcodeFormat format) {
        return format == BarcodeFormat.QR_CODE ? DEFAULT_QR_SIZE : DEFAULT_CODE128_HEIGHT;
    }

    private BufferedImage toBufferedImage(BitMatrix matrix) {
        int width = matrix.getWidth();
        int height = matrix.getHeight();
        BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                image.setRGB(x, y, matrix.get(x, y) ? 0x000000 : 0xFFFFFF);
            }
        }
        return image;
    }
}
