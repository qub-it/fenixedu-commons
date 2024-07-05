package org.fenixedu.commons.spreadsheet;

import java.io.Serializable;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public class ExcelStyle implements Serializable {

    private static final long serialVersionUID = 6778686809629990612L;

    private HSSFCellStyle titleStyle;

    private HSSFCellStyle headerStyle;

    private HSSFCellStyle verticalHeaderStyle;

    private HSSFCellStyle stringStyle;

    private HSSFCellStyle doubleStyle;

    private HSSFCellStyle doubleNegativeStyle;

    private HSSFCellStyle integerStyle;

    private HSSFCellStyle labelStyle;

    private HSSFCellStyle valueStyle;

    private HSSFCellStyle redValueStyle;

    public ExcelStyle(HSSFWorkbook wb) {
        setTitleStyle(wb);
        setHeaderStyle(wb);
        setVerticalHeaderStyle(wb);
        setStringStyle(wb);
        setDoubleStyle(wb);
        setDoubleNegativeStyle(wb);
        setIntegerStyle(wb);
        setLabelStyle(wb);
        setValueStyle(wb);
        setRedValueStyle(wb);
    }

    private void setTitleStyle(HSSFWorkbook wb) {
        HSSFCellStyle style = wb.createCellStyle();
        HSSFFont font = wb.createFont();
        font.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        font.setBold(true);
        font.setFontHeightInPoints((short) 10);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        titleStyle = style;
    }

    private void setHeaderStyle(HSSFWorkbook wb) {
        HSSFCellStyle style = wb.createCellStyle();
        HSSFFont font = wb.createFont();
        font.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        font.setBold(true);
        font.setFontHeightInPoints((short) 8);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillForegroundColor(HSSFColor.HSSFColorPredefined.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        headerStyle = style;
    }

    private void setVerticalHeaderStyle(HSSFWorkbook wb) {
        verticalHeaderStyle = wb.createCellStyle();
        HSSFFont font = wb.createFont();
        font.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        font.setBold(true);
        font.setFontHeightInPoints((short) 8);
        verticalHeaderStyle.setFont(font);
        verticalHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        verticalHeaderStyle.setFillForegroundColor(HSSFColor.HSSFColorPredefined.GREY_25_PERCENT.getIndex());
        verticalHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        verticalHeaderStyle.setBorderLeft(BorderStyle.THIN);
        verticalHeaderStyle.setBorderRight(BorderStyle.THIN);
        verticalHeaderStyle.setBorderBottom(BorderStyle.THIN);
        verticalHeaderStyle.setBorderTop(BorderStyle.THIN);
        verticalHeaderStyle.setRotation((short) 90);
    }

    private void setStringStyle(HSSFWorkbook wb) {
        HSSFCellStyle style = wb.createCellStyle();
        HSSFFont font = wb.createFont();
        font.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        font.setFontHeightInPoints((short) 8);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        stringStyle = style;
    }

    private void setDoubleStyle(HSSFWorkbook wb) {
        HSSFCellStyle style = wb.createCellStyle();
        HSSFFont font = wb.createFont();
        font.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        font.setFontHeightInPoints((short) 8);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
        doubleStyle = style;
    }

    private void setDoubleNegativeStyle(HSSFWorkbook wb) {
        HSSFCellStyle style = wb.createCellStyle();
        HSSFFont font = wb.createFont();
        font.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        font.setFontHeightInPoints((short) 8);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
        font.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
        doubleNegativeStyle = style;
    }

    private void setIntegerStyle(HSSFWorkbook wb) {
        HSSFCellStyle style = wb.createCellStyle();
        HSSFFont font = wb.createFont();
        font.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        font.setFontHeightInPoints((short) 8);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setDataFormat(wb.createDataFormat().getFormat("0"));
        integerStyle = style;
    }

    private void setLabelStyle(HSSFWorkbook wb) {
        HSSFCellStyle style = wb.createCellStyle();
        HSSFFont font = wb.createFont();
        font.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        font.setBold(true);
        font.setFontHeightInPoints((short) 8);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.LEFT);
        labelStyle = style;
    }

    private void setValueStyle(HSSFWorkbook wb) {
        HSSFCellStyle style = wb.createCellStyle();
        HSSFFont font = wb.createFont();
        font.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        font.setFontHeightInPoints((short) 8);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setWrapText(true);
        valueStyle = style;
    }

    private void setRedValueStyle(HSSFWorkbook wb) {
        HSSFCellStyle style = wb.createCellStyle();
        HSSFFont font = wb.createFont();
        font.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
        font.setFontHeightInPoints((short) 8);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setWrapText(true);
        redValueStyle = style;
    }

    public HSSFCellStyle getDoubleNegativeStyle() {
        return doubleNegativeStyle;
    }

    public HSSFCellStyle getDoubleStyle() {
        return doubleStyle;
    }

    public HSSFCellStyle getHeaderStyle() {
        return headerStyle;
    }

    public HSSFCellStyle getIntegerStyle() {
        return integerStyle;
    }

    public HSSFCellStyle getLabelStyle() {
        return labelStyle;
    }

    public HSSFCellStyle getStringStyle() {
        return stringStyle;
    }

    public HSSFCellStyle getTitleStyle() {
        return titleStyle;
    }

    public HSSFCellStyle getValueStyle() {
        return valueStyle;
    }

    public HSSFCellStyle getRedValueStyle() {
        return redValueStyle;
    }

    public HSSFCellStyle getVerticalHeaderStyle() {
        return verticalHeaderStyle;
    }
}
