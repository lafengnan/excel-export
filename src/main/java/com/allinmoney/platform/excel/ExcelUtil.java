package com.allinmoney.platform.excel;

import com.allinmoney.platform.annotation.ExcelAttribute;
import com.allinmoney.platform.annotation.Translate;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.io.IOException;
import java.io.OutputStream;
import java.io.Serializable;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static com.allinmoney.platform.excel.ExcelSheet.MAX_ROW;

/**
 * Created by chris on 16/4/27.
 *
 * @param <T> the type parameter
 */
public class ExcelUtil<T> implements Serializable {

    private static final long serialVersionId = 551970754610248636L;

    private static final Logger logger = Logger.getLogger(ExcelUtil.class);

    private static final short HEADER_FONT_HEIGHT = 14;
    private static final short CONTENT_FONT_HEIGHT = 12;
    private static final int DEFAULT_DELIMITER = 5;

    private Class<T> cls;

    private Class<?> view;

    private int delimiter;

    /**
     * Instantiates a new Excel util.
     *
     * @param cls the cls of data source type
     */
    public ExcelUtil(Class<T> cls) {
        this.cls = cls;
        this.view = null;
        this.delimiter = DEFAULT_DELIMITER;
    }

    /**
     * Sets data view.
     *
     * @param excelView the excel view
     */
    public void setDataView(Class<?> excelView) {
        this.view = excelView;
    }

    /**
     * Sets delimiter.
     *
     * @param delimiter the delimiter
     */
    public void setDelimiter(int delimiter) {
        this.delimiter = delimiter;
    }

    /**
     * Export data list. If @param withSuperFields is true, exportation will not
     * only export given current model's annotated fields, but also export the annotated
     * fields of current model's super class.
     *
     * @param dataList        the data list to export
     * @param sheetName       the sheet name to export
     * @param os              the OutputStream for I/O
     * @param withSuperFields identifies if to export the annotated fields of element's super class
     * @return the boolean
     */
    public boolean exportDataList(List<T> dataList, String sheetName, OutputStream os, boolean withSuperFields) {
        return exportDataList(dataList, sheetName, os, null, withSuperFields);
    }

    /**
     * Export data list with given date format. @since 1.0.5 the new attribute "format"
     * was introduced to the annotation of ExcelAttribute. User can pass null to use the
     * format of annotation.
     *
     * @param dataList  the data list to export
     * @param sheetName the sheet name to export
     * @param os        the OutputStream for I/O
     * @param dateFmt   the date fmt for date fields of data source, if null is passed in,
     *                  the default fmt "yyyy-MM-dd HH:mm:ss" will be used which is defined
     *                  in the annotation with format attribute
     * @return the boolean
     */
    public boolean exportDataList(List<T> dataList, String sheetName, OutputStream os, String dateFmt) {
        return exportDataList(dataList, sheetName, os, dateFmt, true);
    }

    /**
     * Export data list with date format and withSuperFields flag. This method is the low level method
     * to export data. User can use this method explicitly or use other two overloaded methods simply.
     *
     * @param dataList        the data list to export
     * @param sheetName       the sheet name to export
     * @param os              the OutputStream for I/O
     * @param dateFmt         the date fmt for date fields of data source, if null is passed in,
     *                        the default fmt "yyyy-MM-dd HH:mm:ss" will be used which is defined
     *                        in the annotation with format attribute
     * @param withSuperFields identifies if to export the annotated fields of element's super class
     * @return the boolean
     */
    public boolean exportDataList(List<T> dataList, String sheetName, OutputStream os, String dateFmt, boolean withSuperFields) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        int sheets = dataList.size() == 0?1:
                    dataList.size() % MAX_ROW == 0?dataList.size()/MAX_ROW:dataList.size()/MAX_ROW + 1;

        List<Field> annotatedFields = getAnnotatedFields(withSuperFields);
        for (int idx = 0; idx < sheets; idx++) {
            ExcelSheet sheet = new ExcelSheet(workbook.createSheet(sheetName + idx));
            sheet.setWorkbook(workbook)
                    .initStylesAndFonts()
                    .createHeaders(annotatedFields)
                    .fillInContent(annotatedFields, dataList, false, idx, dateFmt)
                    .addSummary(annotatedFields);
        }
        flushWorkbook(workbook, os);
        return true;
    }

    /**
     * Export multiple data list. This method is used to export multiple data source into excel
     * file. Different data source will be exported to different sheets.
     *
     * @param sheetName       the sheet name to export
     * @param withSuperFields identifies if to export the annotated fields of element's super class
     * @param os              OutputStream for I/O
     * @param dataList        the data list array to export
     * @return the boolean
     */
    public boolean exportMultipleDataList(String sheetName, boolean withSuperFields, OutputStream os, List<?>... dataList) {
        int sheetNo = 0;
        HSSFWorkbook workbook = new HSSFWorkbook();
        for (List list : dataList) {
            if (list.isEmpty())
                continue;

            Class<?> clz = list.get(0).getClass();
            List<Field> annotatedFields = getAnnotatedFields(clz, withSuperFields);

            ExcelSheet sheet = new ExcelSheet(workbook.createSheet(sheetName + sheetNo));
            sheet.setWorkbook(workbook)
                    .initStylesAndFonts()
                    .createHeaders(annotatedFields)
                    .fillInContent(annotatedFields, list, true, sheetNo, null)
                    .addSummary(annotatedFields);
            sheetNo++;
        }
        flushWorkbook(workbook, os);
        return true;
    }

    /**
     * Export multiple data source within same sheet.
     *
     * @param sheetName the sheet name to export
     * @param os        OutputStream for I/O
     * @param dataList  the data list array to export
     * @return the boolean
     */
    public boolean exportDataList(String sheetName, OutputStream os, List<?>... dataList) {
        HSSFWorkbook workbook = new HSSFWorkbook();

        HSSFSheet sheet = workbook.createSheet(sheetName);
        HSSFRow row;
        HSSFCell headerCell;
        HSSFCell contentCell;

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

        int line = 0;
        int first = 0;
        try {
            for (List list:dataList) {
                if (list.size() == 0) {
                    continue;
                }

                Class clazz = list.get(0).getClass();// 获取集合中的对象类型
                Field[] fields = clazz.getDeclaredFields();// 获取他的字段数组

                List<Field> validFields = new LinkedList<>(); // fields annotated with ExcelAttribute
                for (Field f: fields) {
                    if (f.isAnnotationPresent(ExcelAttribute.class)) {
                        ExcelAttribute attr = f.getAnnotation(ExcelAttribute.class);
                        boolean match = false;

                        if (this.view == null) {
                            match = true;
                        } else {
                            for (Class<?> v:attr.groups()) {
                                if (v.equals(this.view)) {
                                    match = true;
                                    break;
                                }
                            }
                        }
                        if (match) {
                            validFields.add(f);
                        }
                    }
                }

                row = sheet.createRow(line);
                line++;
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
                    sheet.autoSizeColumn(col);
                }

                contentFont = workbook.createFont();
                first = line;
                for (int i = 0; i < list.size(); i++) {
                    row = sheet.createRow(line);
                    line++;
                    Object data = list.get(i);
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
                                    SimpleDateFormat sdf = new SimpleDateFormat(attr.format());
                                    txtValue = sdf.format(date);
                                } else if(field.get(data) instanceof BigDecimal) {
                                    BigDecimal big = (BigDecimal) field.get(data);
                                    txtValue = big.setScale(2, BigDecimal.ROUND_HALF_EVEN).toString();
                                } else {
                                    txtValue = field.get(data) == null?"":field.get(data).toString();
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
                                e.printStackTrace();
                                logger.debug(e);
                            }
                            sheet.autoSizeColumn(col);
                        }
                    }
                }

                // create summary row
                HSSFRow sumRow = sheet.createRow(line);
                line++;
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
                        for (int j = first; j < lastRowNum; j++) {
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
                        sumCell.setCellValue(new HSSFRichTextString("合计: " + sum.setScale(2, BigDecimal.ROUND_HALF_EVEN).toString()));
                        sheet.autoSizeColumn(col);
                    }
                }
                line=line+delimiter;
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
     * Gets excel col.
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

    /**
     * Sets hssf prompt.
     *
     * @param sheet         the sheet
     * @param promptTitle   the prompt title
     * @param promptContent the prompt content
     * @param firstRow      the first row
     * @param endRow        the end row
     * @param firstCol      the first col
     * @param endCol        the end col
     * @return the hssf prompt
     */
    public static HSSFSheet setHSSFPrompt(HSSFSheet sheet, String promptTitle, String promptContent, int firstRow, int endRow, int firstCol, int endCol) {
        DVConstraint constraint = DVConstraint.createCustomFormulaConstraint("DD1");

        CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, firstCol, endCol);
        HSSFDataValidation validation = new HSSFDataValidation(regions, constraint);
        validation.createPromptBox(promptTitle, promptContent);
        sheet.addValidationData(validation);
        return sheet;
    }

    /**
     * Sets hssf validation.
     *
     * @param sheet    the sheet
     * @param txtList  the txt list
     * @param firstRow the first row
     * @param endRow   the end row
     * @param firstCol the first col
     * @param endCol   the end col
     * @return the hssf validation
     */
    public static HSSFSheet setHSSFValidation(HSSFSheet sheet, String[] txtList, int firstRow, int endRow, int firstCol, int endCol) {
        DVConstraint constraint = DVConstraint.createExplicitListConstraint(txtList);
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, endRow, firstCol, endCol);
        HSSFDataValidation validationList = new HSSFDataValidation(regions, constraint);
        sheet.addValidationData(validationList);
        return sheet;
    }

    private Optional<Field[]> getFieldsOfSuperClass(Class<T> clz) {
        Optional<Field[]> fields = Optional.empty();
        if (clz.getClass().getSuperclass() != null) {
            fields = Optional.of(clz.getClass().getSuperclass().getDeclaredFields());
        }
        return fields;
    }

    private List<Field> getAnnotatedFields(boolean superFlag) {
        return getAnnotatedFields(cls, superFlag);
    }

    private List<Field> getAnnotatedFields(Class<?> clz, boolean superFlag) {
        List<Field> fields = new LinkedList<>();
        List<Field> annotatedFields = new LinkedList<>();
        if (superFlag && clz.getSuperclass() != null) {
            fields.addAll(Arrays.asList(clz.getSuperclass().getDeclaredFields()));
        }
        fields.addAll(Arrays.asList(clz.getDeclaredFields()));

        fields.stream()
                .filter(f->f.isAnnotationPresent(ExcelAttribute.class))
                .forEach(f->{
                    boolean match = false;
                    ExcelAttribute attr = f.getAnnotation(ExcelAttribute.class);
                    if (this.view == null) {
                        match = true;
                    } else {
                        for (Class<?> v : attr.groups()) {
                            if (v.equals(this.view)) {
                                match = true;
                                break;
                            }
                        }
                    }
                    if (match) {
                        annotatedFields.add(f);
                    }
                });

        return annotatedFields;
    }

    private void flushWorkbook(HSSFWorkbook workbook, OutputStream os) throws RuntimeException {
        try {
            os.flush();
            workbook.write(os);
            os.close();
        } catch (IOException e) {
            e.printStackTrace();
            logger.info(e.getMessage());
            throw new ExcelException(e.getMessage());
        }
    }

}
