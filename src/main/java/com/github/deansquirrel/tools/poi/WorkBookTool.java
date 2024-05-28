package com.github.deansquirrel.tools.poi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.List;

public class WorkBookTool {

    private WorkBookTool(){};

    /**
     * 获取字体
     * @param book 表格对象
     * @return 字体
     */
    public static Font getFont(Workbook book) {
        if(book == null) {
            return null;
        }
        Font font = book.createFont();
        font.setFontName("Calibri");
        font.setBold(false);
        font.setFontHeightInPoints((short) 11);
        return font;
    }

    /**
     * 获取日期时间类型的单元格样式
     * @param book 表格对象
     * @param font 字体文件
     * @return 单元格样式
     */
    public static CellStyle getDateTimeCellStyle(Workbook book, Font font) {
        if(book == null) {
            return null;
        }
        CreationHelper creationHelper = book.getCreationHelper();
        CellStyle cellStyle = book.createCellStyle();
        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));
        if(font != null) {
            cellStyle.setFont(font);
        }
        return cellStyle;
    }

    /**
     * 获取日期时间类型的单元格样式
     * @param book 表格对象
     * @return 单元格样式
     */
    public static CellStyle getDateTimeCellStyle(Workbook book) {
        return getDateTimeCellStyle(book, getFont(book));
    }

    /**
     * 获取日期类型的单元格样式
     * @param book 表格对象
     * @param font 字体
     * @return 单元格样式
     */
    public static CellStyle getDateCellStyle(Workbook book, Font font) {
        if(book == null) {
            return null;
        }
        CreationHelper creationHelper = book.getCreationHelper();
        CellStyle cellStyle = book.createCellStyle();
        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-mm-dd"));
        if(font != null) {
            cellStyle.setFont(font);
        }
        return cellStyle;
    }

    /**
     * 获取日期类型的单元格样式
     * @param book 表格对象
     * @return 单元格样式
     */
    public static CellStyle getDateCellStyle(Workbook book) {
        if(book == null) {
            return null;
        }
        return getDateCellStyle(book, getFont(book));
    }

    /**
     * 获取单元格样式
     * @param book 表格对象
     * @param font 字体
     * @return 单元格样式
     */
    public static CellStyle getCellStyle(Workbook book, Font font) {
        if(book == null) {
            return null;
        }
        CellStyle cellStyle = book.createCellStyle();
        cellStyle.setFont(font);
        return cellStyle;
    }

    /**
     * 获取单元格样式
     * @param book 表格对象
     * @return 单元格样式
     */
    public static CellStyle getCellStyle(Workbook book) {
        if(book == null) {
            return null;
        }
        CellStyle cellStyle = book.createCellStyle();
        cellStyle.setFont(getFont(book));
        return cellStyle;
    }

    private static final SimpleDateFormat checkHMS = new SimpleDateFormat("HHmmss");
    private static final String DATE_ZERO = "000000";

    private static CellStyle getCellStyle(Object d,
                                          CellStyle dateStyle,
                                          CellStyle dateTimeStyle,
                                          CellStyle cellStyle) {
        if(d == null) {
            return cellStyle;
        }
        if(d instanceof Date) {
            return DATE_ZERO.equals(checkHMS.format(d)) ? dateStyle : dateTimeStyle;
        } else {
            return cellStyle;
        }
    }

    private static void createRow(Sheet sheet, List<Object> data, List<CellStyle> cellStyleList, int index,
                                  CellStyle dateStyle,
                                  CellStyle dateTimeStyle,
                                  CellStyle cellStyle) {
        if(sheet == null || data == null || data.isEmpty()) {
            return;
        }
        Row row = sheet.createRow(Math.max(index, 0));
        for(int i = 0 ; i < data.size(); i++) {
            Object d = data.get(i);
            if(d == null) {
                continue;
            }
            CellStyle currCellStyle = (cellStyleList == null || cellStyleList.isEmpty() || cellStyleList.get(i) == null)
                    ? getCellStyle(d, dateStyle, dateTimeStyle, cellStyle) : cellStyleList.get(i);
            Cell cell = row.createCell(i);
            if(d instanceof Date) {
                cell.setCellValue((Date)d);
            } else {
                RichTextString richTextString = new XSSFRichTextString(String.valueOf(d));
                cell.setCellValue(richTextString);
            }
            cell.setCellStyle(currCellStyle);
        }
    }

    public static SXSSFWorkbook getSXSSFWorkBook(List<WorkTableData> list) {
        return getSXSSFWorkBook(list, null);
    }

    public static SXSSFWorkbook getSXSSFWorkBook(List<WorkTableData> list, List<CellStyle> cellStyleList) {
        SXSSFWorkbook book = new SXSSFWorkbook();
        if(list == null || list.isEmpty()) {
            book.createSheet("Sheet1");
        } else {
            Font font = getFont(book);
            CellStyle cellDateTimeStyle = getDateTimeCellStyle(book, font);
            CellStyle cellDateStyle = getDateCellStyle(book, font);
            CellStyle cellStyle = getCellStyle(book, font);
            for (WorkTableData table : list) {
                SXSSFSheet sheet = book.createSheet(table.getName());
                List<String> title = table.getTitle();
                if (title != null && !title.isEmpty()) {
                    createRow(sheet, new ArrayList<>(title), null, 0,
                            cellDateStyle, cellDateTimeStyle, cellStyle);
                }
                List<List<Object>> dataList = table.getRows();
                if (dataList != null && !dataList.isEmpty()) {
                    for (int i = 0; i < dataList.size(); i++) {
                        List<CellStyle> currCellStyleList = (cellStyleList == null
                                || cellStyleList.isEmpty()) ?
                                null : cellStyleList;
                        createRow(sheet, dataList.get(i), currCellStyleList, i + 1,
                                cellDateStyle, cellDateTimeStyle, cellStyle);
                    }
                }
            }
        }
        return book;
    }

    public static <T> WorkTableData getXSSFWorkTable(String name, List<T> list, IDataMapper<T> iDataMapper) {
        List<List<Object>> rows = new ArrayList<>();
        for(T data : list) {
            rows.add(iDataMapper.getRowData(data));
        }
        return WorkTableData.builder(name)
                .initTitle(iDataMapper.getTitleList())
                .initRows(rows);
    }

    public static void saveSXSSFWorkbook(String base, String fileName, SXSSFWorkbook f) throws IOException {
        String fullPath = base + ((base == null || base.isEmpty()) ? "" : "/") + fileName;
        try (FileOutputStream fs = new FileOutputStream(fullPath)) {
            f.write(fs);
            fs.flush();
        }
    }

}
