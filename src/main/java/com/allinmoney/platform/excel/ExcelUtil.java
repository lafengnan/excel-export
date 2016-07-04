package com.allinmoney.platform.excel;

import com.allinmoney.platform.annotation.ExcelAttribute;
import com.allinmoney.platform.annotation.Translate;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.io.OutputStream;
import java.io.Serializable;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by chris on 16/4/27.
 */
public class ExcelUtil<T> implements Serializable {

    private static final long serialVersionId = 551970754610248636L;

    private static final Logger logger = Logger.getLogger(ExcelUtil.class);

    private static final int MAX_ROW = 65535;
    private static final short HEADER_FONT_HEIGHT = 14;
    private static final short CONTENT_FONT_HEIGHT = 12;


    private Class<T> cls;

    public ExcelUtil(Class<T> cls) {
        this.cls = cls;
    }

    public boolean exportDataList(List<T> dataList, String sheetName, OutputStream os, String fmt) {

        try {
            Field[] fields = cls.getDeclaredFields(); // All fields of one java bean
            List<Field> validFields = new LinkedList<>(); // fields annotated with ExcelAttribute
            for (Field f: fields) {
                if (f.isAnnotationPresent(ExcelAttribute.class)) {
                    validFields.add(f);
                }
            }

            HSSFWorkbook workbook = new HSSFWorkbook();
            double sheetNumbers = Math.ceil(dataList != null?dataList.size():1/MAX_ROW);
            int sheets = dataList.size() == 0?1:
                    dataList.size() % MAX_ROW == 0?dataList.size()/MAX_ROW:dataList.size()/MAX_ROW + 1;
            for (int idx = 0; idx < sheets; idx++) {
            //for (int idx = 0; idx < sheetNumbers; idx++) {
                HSSFSheet sheet = workbook.createSheet(sheetName + idx);
                HSSFRow row;
                HSSFCell headerCell;
                HSSFCell contentCell;
                row = sheet.createRow(0);

                // set style for normal cell
                HSSFCellStyle headerCellStyle = workbook.createCellStyle();
                headerCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);

                // Font
                HSSFFont headerFont = workbook.createFont();
                headerFont.setFontName("Arail narrow");
                headerFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
                headerFont.setFontHeightInPoints(HEADER_FONT_HEIGHT);

                // set style for content
                HSSFCellStyle contentCellStyle = workbook.createCellStyle();

                // Font
                HSSFFont contentFont = workbook.createFont();
                contentFont.setFontName("Arail narrow");
                contentFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
                contentFont.setFontHeightInPoints(CONTENT_FONT_HEIGHT);

                // Special style for mark
                HSSFCellStyle markCellStyle = workbook.createCellStyle();

                // Header Font
                HSSFFont markHeaderFont = workbook.createFont();
                markHeaderFont.setFontName("Arail narrow");
                markHeaderFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
                markHeaderFont.setFontHeightInPoints(HEADER_FONT_HEIGHT);

                // Content Font
                HSSFFont markContentFont = workbook.createFont();
                markContentFont.setFontName("Arail narrow");
                markContentFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
                markContentFont.setFontHeightInPoints(CONTENT_FONT_HEIGHT);

                // create headers
                for (int i = 0; i < validFields.size(); i++) {
                    Field field = validFields.get(i);
                    ExcelAttribute attr = field.getAnnotation(ExcelAttribute.class);
                    int col = i;

                    if (StringUtils.isNotBlank(attr.column())) {
                        col = getExcelCol(attr.column());
                    }

                    // create columns
                    headerCell = row.createCell(col);
                    if (attr.isMark()) {
                        markHeaderFont.setColor(HSSFColor.RED.index);
                        markCellStyle.setFont(markHeaderFont);
                        headerCell.setCellStyle(markCellStyle);
                    } else {
                        headerFont.setColor(HSSFColor.BLACK.index);
                        headerCellStyle.setFont(headerFont);
                        headerCell.setCellStyle(headerCellStyle);
                    }

//                    sheet.setColumnWidth(i, (int)((attr.name().getBytes().length <= 4?4:attr.name().getBytes().length * 1.5 * 256)));
                    headerCell.setCellType(HSSFCell.CELL_TYPE_STRING);
                    headerCell.setCellValue(attr.title());

                    if (StringUtils.isNotBlank(attr.prompt())) {
                        setHSSFPrompt(sheet, "", attr.prompt(), 1, 100, col, col);
                    }

                    if (attr.combo().length > 0) {
                        setHSSFValidation(sheet, attr.combo(), 1, 100, col, col);
                    }
                }

                // fill in content
                contentFont = workbook.createFont();
                int startNo = idx * MAX_ROW;
                int endNo = Math.min(startNo + MAX_ROW, dataList == null?0:dataList.size());

                for (int i = startNo; i < endNo; i++) {
                    row = sheet.createRow(i + 1 - startNo);
                    T data = (T)dataList.get(i);
                    for (int j = 0; j < validFields.size(); j++) {
                        Field field = validFields.get(j);
                        field.setAccessible(true);
                        ExcelAttribute attr = field.getAnnotation(ExcelAttribute.class);

                        int col = j;
                        if (StringUtils.isNotBlank(attr.column())) {
                            col = getExcelCol(attr.column());
                        }

                        if (attr.isExport()) {
                            contentCell = row.createCell(col);
                            if (attr.isMark()) {
                                markContentFont.setColor(HSSFFont.COLOR_RED);
                                markCellStyle.setFont(contentFont);
                                contentCell.setCellStyle(markCellStyle);
                            } else {
                                contentFont.setColor(HSSFFont.COLOR_NORMAL);
                                contentCellStyle.setFont(contentFont);
                                contentCell.setCellStyle(contentCellStyle);
                            }

                            try {
                                String txtValue = null;
                                if (field.get(data) instanceof Date) {
                                    Date date = (Date) field.get(data);
                                    SimpleDateFormat sdf = new SimpleDateFormat(fmt);
                                    txtValue = sdf.format(date);
                                } else {
                                    txtValue = field.get(data) == null?"null":field.get(data).toString();
                                }

                                Map<String, String> map = new HashMap<>(); // translate map
                                if (attr.translate().length > 0) {
                                    Translate[] translates = attr.translate();
                                    for (int ix = 0; ix < translates.length; ix++) {
                                        map.put(translates[ix].key(), translates[ix].value());
                                    }
                                }

                                Pattern p = Pattern.compile("^//d+(//.//d+)?$");
                                Matcher matcher = p.matcher(txtValue);
                                if (matcher.matches())
                                {
                                    if (map.containsKey(txtValue)) {
                                        contentCell.setCellValue(Double.parseDouble(map.get(txtValue)));
                                    } else {
                                        contentCell.setCellValue(Double.parseDouble(txtValue));
                                    }
                                } else {
                                    contentCell.setCellValue(map.containsKey(txtValue)?map.get(txtValue):txtValue);
                                }

                            } catch (IllegalAccessException e) {
                                logger.debug(e);
                            }
                        }
                    }
                }

                // create summary row
                HSSFRow sumRow = sheet.createRow((short)(sheet.getLastRowNum() + 1));
                for (int i = 0; i < validFields.size(); i++) {
                    Field field  = validFields.get(i);
                    ExcelAttribute attr = field.getAnnotation(ExcelAttribute.class);
                    if (attr.isSum()) {
                        int col = i;
                        if (StringUtils.isNotBlank(attr.column())) {
                            col = getExcelCol(attr.column());
                        }
                        BigDecimal sum = BigDecimal.ZERO;
                        int lastRowNum = sheet.getLastRowNum();
                        for (int j = 1; j < lastRowNum; j++) {
                            HSSFRow idxRow = sheet.getRow(j);
                            if (idxRow != null) {
                                HSSFCell idxCell = idxRow.getCell(col);
                                if (idxCell != null &&
                                        idxCell.getCellType() == HSSFCell.CELL_TYPE_STRING &&
                                        NumberUtils.isNumber(idxCell.getStringCellValue())) {
                                    sum = sum.add(BigDecimal.valueOf(Double.valueOf(idxCell.getStringCellValue())));
                                }
                            }
                        }
                        HSSFCell sumCell = sumRow.createCell(col);
                        sumCell.setCellValue(new HSSFRichTextString("合计: " + sum));
                    }
                }

            }
            os.flush();
            workbook.write(os);
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
            throw new ExcelException(e.getMessage());
        }
        return true;
    }

    /**
     *
     * @param col column name A, B, C...
     * @return int value according to A, B, C...
     */
    public static int getExcelCol(String col) {
        col = col.toUpperCase();
        int count = -1;
        char[] cs = col.toCharArray();
        for (int i = 0; i < cs.length; i++) {
            count += (cs[i] - 64) * Math.pow(26, cs.length - 1 - i);
        }

        return count;
    }

    public static HSSFSheet setHSSFPrompt(HSSFSheet sheet, String promptTitle, String promptContent, int firstRow, int endRow, int firstCol, int endCol) {
        DVConstraint constraint = DVConstraint.createCustomFormulaConstraint("DD1");

        CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, firstCol, endCol);
        HSSFDataValidation validation = new HSSFDataValidation(regions, constraint);
        validation.createPromptBox(promptTitle, promptContent);
        sheet.addValidationData(validation);
        return sheet;
    }

    public static HSSFSheet setHSSFValidation(HSSFSheet sheet, String[] txtList, int firstRow, int endRow, int firstCol, int endCol) {
        DVConstraint constraint = DVConstraint.createExplicitListConstraint(txtList);
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, firstCol, endCol);
        HSSFDataValidation validationList = new HSSFDataValidation(regions, constraint);
        sheet.addValidationData(validationList);
        return sheet;
    }

}
