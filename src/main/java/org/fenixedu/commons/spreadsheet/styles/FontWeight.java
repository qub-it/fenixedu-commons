package org.fenixedu.commons.spreadsheet.styles;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class FontWeight extends CellStyle {

    static final short BOLDWEIGHT_BOLD = 0x2bc;

    private final boolean boldweight;

    public FontWeight(short boldweight) {
        this.boldweight = boldweight == BOLDWEIGHT_BOLD;
    }

    public FontWeight(boolean isBold) {
        this.boldweight = isBold;
    }

    @Override
    protected void appendToStyle(HSSFWorkbook book, HSSFCellStyle style, HSSFFont font) {
        font.setBold(boldweight);
    }

    @Override
    public HSSFCellStyle getStyle(HSSFWorkbook book) {
        HSSFCellStyle style = book.createCellStyle();
        HSSFFont font = book.createFont();
        appendToStyle(book, style, font);
        style.setFont(font);
        return style;
    }

    @Override
    public boolean equals(Object obj) {
        if (obj instanceof FontWeight) {
            FontWeight fontWeight = (FontWeight) obj;
            return boldweight == fontWeight.boldweight;
        }
        return false;
    }

    @Override
    public int hashCode() {
        return boldweight ? BOLDWEIGHT_BOLD : Boolean.hashCode(boldweight);
    }
}
