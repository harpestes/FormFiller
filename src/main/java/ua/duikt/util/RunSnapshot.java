package ua.duikt.util;

import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class RunSnapshot {
    private final String fontFamily;
    private final double fontSize;
    private final boolean bold;
    private final boolean italic;
    private final UnderlinePatterns underline;
    private final String color;

    public RunSnapshot(XWPFRun run) {
        this.fontFamily = run.getFontFamily();
        if(run.getFontSizeAsDouble() != null) {
            this.fontSize = run.getFontSizeAsDouble();
        } else {
            this.fontSize = 11;
        }
        this.bold = run.isBold();
        this.italic = run.isItalic();
        this.underline = run.getUnderline();
        this.color = run.getColor();
    }

    public void applyTo(XWPFRun run) {
        if (fontFamily != null) run.setFontFamily(fontFamily);
        if (fontSize > 0) run.setFontSize(fontSize);
        run.setBold(bold);
        run.setItalic(italic);
        run.setUnderline(underline);
        if (color != null) run.setColor(color);
    }
}

