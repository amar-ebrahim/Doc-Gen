package com.bawdocgen.barcode;

public class BarcodeImage {
    private final byte[] pngBytes;
    private final int width;
    private final int height;

    public BarcodeImage(byte[] pngBytes, int width, int height) {
        this.pngBytes = pngBytes;
        this.width = width;
        this.height = height;
    }

    public byte[] getPngBytes() {
        return pngBytes;
    }

    public int getWidth() {
        return width;
    }

    public int getHeight() {
        return height;
    }
}
