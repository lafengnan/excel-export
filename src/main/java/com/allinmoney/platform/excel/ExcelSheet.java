package com.allinmoney.platform.excel;

import com.allinmoney.platform.annotation.ExcelAttribute;
import com.allinmoney.platform.annotation.Translate;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by chris on 2017/1/17.
 */
public class ExcelSheet {
    private HSSFSheet sheet;
    private HSSFFont headerFont;
    private HSSFFont contentFont;
    private HSSFFont markHeaderFont;
    private HSSFFont markContentFont;
    private HSSFCellStyle headerCellStyle;
    private HSSFCellStyle markHeaderCellStyle;
    private HSSFCellStyle cellStyle;
    private HSSFCellStyle markCellStyle;
    private HSSFWorkbook workbook;

    /**
     * The constant MAX_ROW.
     */
    public static final int MAX_ROW = 1_048_576;
    private static final short HEADER_FONT_HEIGHT = 14;
    private static final short CONTENT_FONT_HEIGHT = 12;
    private static final String FONT_NAME = "Arail narrow";
    private static final Logger logger = Logger.getLogger(ExcelSheet.class);

    /**
     * Instantiates a new Excel sheet.
     *
     * @param sheet the sheet
     */
    public ExcelSheet(HSSFSheet sheet) {
        this.sheet = sheet;
    }

    /**
     * Sets sheet.
     *
     * @param sheet the sheet
     */
    public void setSheet(HSSFSheet sheet) {
        this.sheet = sheet;
    }


    /**
     * Gets sheet.
     *
     * @return the sheet
     */
    public HSSFSheet getSheet() {
        return sheet;
    }

    /**
     * Gets workbook.
     *
     * @return the workbook
     */
    public HSSFWorkbook getWorkbook() {
        return workbook;
    }

    /**
     * Sets workbook.
     *
     * @param workbook the workbook
     * @return the workbook
     */
    public ExcelSheet setWorkbook(HSSFWorkbook workbook) {
        if (this.workbook == null)
            this.workbook = workbook;
        return this;
    }

    /**
     * Init styles and fonts excel sheet.
     *
     * @return the excel sheet
     */
    public ExcelSheet initStylesAndFonts() {
        headerFont = workbook.createFont();
        contentFont = workbook.createFont();
        markHeaderFont = workbook.createFont();
        markContentFont = workbook.createFont();
        setHeaderFont();
        setContentFont();
        setMarkHeaderFont();
        setMarkContentFont();

        cellStyle = workbook.createCellStyle();
        markCellStyle = workbook.createCellStyle();
        headerCellStyle = workbook.createCellStyle();
        markHeaderCellStyle = workbook.createCellStyle();
        cellStyle.setFont(contentFont);
        markCellStyle.setFont(markContentFont);
        headerCellStyle.setFont(headerFont);
        markHeaderCellStyle.setFont(markHeaderFont);

        // header alignment
        headerCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        markHeaderCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        return this;
    }

    /**
     * Gets cell style.
     *
     * @return the cell style
     */
    public HSSFCellStyle getCellStyle() {
        return cellStyle;
    }

    /**
     * Gets mark cell style.
     *
     * @return the mark cell style
     */
    public HSSFCellStyle getMarkCellStyle() {
        return markCellStyle;
    }

    /**
     * Sets header font.
     */
    public void setHeaderFont() {
        headerFont.setFontName(FONT_NAME);
        headerFont.setColor(HSSFColor.BLACK.index);
        headerFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        headerFont.setFontHeightInPoints(HEADER_FONT_HEIGHT);
    }

    /**
     * Sets content font.
     */
    public void setContentFont() {
        contentFont.setFontName(FONT_NAME);
        contentFont.setColor(HSSFFont.COLOR_NORMAL);
        contentFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        contentFont.setFontHeightInPoints(CONTENT_FONT_HEIGHT);
    }

    /**
     * Sets mark header font.
     */
    public void setMarkHeaderFont() {
        markHeaderFont.setFontName(FONT_NAME);
        markHeaderFont.setColor(HSSFColor.RED.index);
        markHeaderFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        markHeaderFont.setFontHeightInPoints(HEADER_FONT_HEIGHT);
    }

    /**
     * Sets mark content font.
     */
    public void setMarkContentFont() {
        markContentFont.setFontName(FONT_NAME);
        markContentFont.setColor(HSSFFont.COLOR_RED);
        markContentFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        markContentFont.setFontHeightInPoints(CONTENT_FONT_HEIGHT);
    }

    /**
     * Exchange columns int. If some column is set column number explicitly,
     * it may occupy the column that has been filled with contents. For such
     * case, we should exchange it to avoid miss data
     *
     * @param fields the fields
     * @param idx    the idx the field index in fields
     * @return the int
     */
    @Deprecated
    public int exchangeColumns(List<Field> fields, int idx) {
        Field field = fields.get(idx);
        ExcelAttribute attribute = field.getAnnotation(ExcelAttribute.class);
        if (StringUtils.isNotBlank(attribute.column())) {
            int col = ExcelUtil.getExcelCol(attribute.column());
            if (col < idx || col > idx && col < fields.size()) {
                //exchange columns
                Field origin = fields.get(col);
                fields.set(col, field);
                fields.set(idx, origin);
                idx--; // back and again
            }
        }
        return idx;
    }

    /**
     * Create headers excel sheet.
     *
     * @param fields the fields
     * @return the excel sheet
     */
    public ExcelSheet addHeaders(List<Field> fields) {
        HSSFRow row = sheet.createRow(0); // init row
        for (int i = 0; i < fields.size(); i++) {
            Field field = fields.get(i);
            ExcelAttribute attribute = field.getAnnotation(ExcelAttribute.class);
            if (!attribute.isExport())
                continue;

            // create columns
            HSSFCell headerCell = row.createCell(i);
            if (attribute.isMark()) {
                headerCell.setCellStyle(markHeaderCellStyle);
            } else {
                headerCell.setCellStyle(headerCellStyle);
            }

            headerCell.setCellType(HSSFCell.CELL_TYPE_STRING);
            headerCell.setCellValue(attribute.title());

            if (StringUtils.isNotBlank(attribute.prompt())) {
                ExcelUtil.setHSSFPrompt(sheet, "提示", attribute.prompt(), 1, 100, i, i);
            }

            if (attribute.combo().length > 0) {
                ExcelUtil.setHSSFValidation(sheet, attribute.combo(), 1, 100, i, i);
            }
            sheet.autoSizeColumn(i);
        }
        return this;
    }

    /**
     * Fill in content excel sheet.
     *
     * @param fields       the fields
     * @param dataList     the data list
     * @param multipleFlag the multiple flag
     * @param sheetNo      the sheet no
     * @param dateFmt      the date fmt
     * @return the excel sheet
     */
    public ExcelSheet addContent(List<Field> fields, List<?> dataList, boolean multipleFlag, int sheetNo, String dateFmt) {
        int startNo = multipleFlag?0:sheetNo * MAX_ROW;
        int endNo = multipleFlag?dataList.size():Math.min(startNo + MAX_ROW, dataList.size());
        if (dataList.size() < startNo)
            return this;

        HSSFCell contentCell;
        for (int i = startNo; i < endNo; i++) {
            HSSFRow row = sheet.createRow(i + 1 - startNo);
            Object data = dataList.get(i);

            for (int j = 0; j < fields.size(); j++) {
                Field field = fields.get(j);
                field.setAccessible(true);

                ExcelAttribute attribute = field.getAnnotation(ExcelAttribute.class);
                if (!attribute.isExport())
                    continue;

                contentCell = row.createCell(j);
                if (attribute.isMark()) {
                    contentCell.setCellStyle(getMarkCellStyle());
                } else {
                    contentCell.setCellStyle(getCellStyle());
                }

                try {
                    String txtValue = "";
                    if (field.get(data) instanceof Date) {
                        Date date = (Date) field.get(data);
                        SimpleDateFormat sdf = new SimpleDateFormat(dateFmt != null?dateFmt:attribute.format());
                        txtValue = sdf.format(date);
                    } else if (field.get(data) instanceof BigDecimal) {
                        BigDecimal big = (BigDecimal) field.get(data);
                        txtValue = big.setScale(2, BigDecimal.ROUND_HALF_EVEN).toString();
                    } else {
                        if (field.get(data) != null)
                            txtValue = field.get(data).toString();
                    }

                    // translate
                    Map<String, String> map = new HashMap<>();
                    if (attribute.translate().length > 0) {
                        Translate[] translates = attribute.translate();
                        for (Translate translate : translates) {
                            map.put(translate.key(), translate.value());
                        }
                    }

                    // digit number
                    Pattern p = Pattern.compile("^//d+(//.//d+)?$");
                    Matcher matcher = p.matcher(txtValue);
                    if (matcher.matches()) {
                        if (map.containsKey(txtValue)) {
                            contentCell.setCellValue(Double.parseDouble(map.get(txtValue)));
                        } else {
                            contentCell.setCellValue(Double.parseDouble(txtValue));
                        }
                    } else {
                        contentCell.setCellValue(map.getOrDefault(txtValue, txtValue));
                    }
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                } finally {
                    sheet.autoSizeColumn(j);
                }
            }
        }
        return this;
    }

    /**
     * Add summary excel sheet.
     *
     * @param fields the fields
     * @return the excel sheet
     */
    public ExcelSheet addSummary(List<Field> fields) {

        HSSFRow sumRow = sheet.createRow(sheet.getLastRowNum() + 1);
        for (int i = 0; i < fields.size(); i++) {
            Field field = fields.get(i);
            ExcelAttribute attribute = field.getAnnotation(ExcelAttribute.class);
            if (!attribute.isSum())
                continue;

            BigDecimal sum = BigDecimal.ZERO;
            int lastRowNum = sheet.getLastRowNum();
            for (int j = 0; j < lastRowNum; j++) {
                HSSFRow idxRow = sheet.getRow(j);
                if (idxRow != null) {
                    HSSFCell idxCell = idxRow.getCell(i);
                    if (idxCell != null &&
                            idxCell.getCellType() == HSSFCell.CELL_TYPE_STRING &&
                            NumberUtils.isNumber(idxCell.getStringCellValue())) {
                        sum = sum.add(BigDecimal.valueOf(Double.valueOf(idxCell.getStringCellValue())));
                    }
                }
            }
            HSSFCell sumCell = sumRow.createCell(i);
            sumCell.setCellValue(new HSSFRichTextString("合计: " + sum.setScale(2, BigDecimal.ROUND_HALF_EVEN).toString()));
            sheet.autoSizeColumn(i);
        }
        return this;
    }
}
