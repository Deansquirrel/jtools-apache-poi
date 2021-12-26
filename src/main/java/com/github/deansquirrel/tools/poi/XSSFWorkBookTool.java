package com.github.deansquirrel.tools.poi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class XSSFWorkBookTool {

    private XSSFWorkBookTool(){};

    public static XSSFWorkbook getXSSFWorkBook(List<XSSFWorkTable> list) {
        XSSFWorkbook book = new XSSFWorkbook();
        if(list == null || list.size() <= 0) {
            book.createSheet("Sheet1");
            return book;
        }
        //字体
        Font font = book.createFont();
        font.setFontName("Calibri");
        font.setBold(false);
        font.setFontHeightInPoints((short) 11);
        //日期格式
        CreationHelper creationHelper = book.getCreationHelper();
        CellStyle cellDateStyle = book.createCellStyle();
        cellDateStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));
        cellDateStyle.setFont(font);
        //日期格式
        CellStyle cellStyle = book.createCellStyle();
        cellStyle.setFont(font);

        for(XSSFWorkTable table : list) {
            XSSFSheet sheet = book.createSheet(table.getName());
            List<String> title = table.getTitle();
            int ret = 0;
            if(title != null && title.size() > 0) {
                ret = 1;
                Row titleRow = sheet.createRow(0);
                for(int i = 0; i < title.size(); i++) {
                    if(title.get(i) == null) continue;
                    Cell cell = titleRow.createCell(i);
                    RichTextString richTextString = new XSSFRichTextString(title.get(i));
                    richTextString.applyFont(font);
                    cell.setCellValue(richTextString);
                    cell.setCellStyle(cellStyle);
                }
            }
            if(table.getRows() != null && table.getRows().size() > 0) {
                for(int i = 0; i < table.getRows().size(); i++) {
                    Row dataRow = sheet.createRow(i + ret);
                    List<Object> rowData = table.getRows().get(i);
                    if(rowData == null) continue;
                    for(int j = 0; j < rowData.size(); j++) {
                        Object cellData = rowData.get(j);
                        if(cellData == null) continue;
                        Cell cell = dataRow.createCell(j);
                        if(cellData instanceof Date) {
                            cell.setCellValue((Date)cellData);
                            cell.setCellStyle(cellDateStyle);
                        } else {
                            RichTextString richTextString = new XSSFRichTextString(String.valueOf(cellData));
                            richTextString.applyFont(font);
                            cell.setCellValue(richTextString);
                            cell.setCellStyle(cellStyle);
                        }
                    }
                }
                ret = ret + table.getRows().size();
            }
        }
        return book;
    }

    public static <T> XSSFWorkTable getXSSFWorkTable(String name, List<T> list, IDataMapper<T> iDataMapper) {
        List<List<Object>> rows = new ArrayList<>();
        for(T data : list) {
            rows.add(iDataMapper.getRowData(data));
        }
        return XSSFWorkTable.builder(name)
                .initTitle(iDataMapper.getTitleList())
                .initRows(rows);
    }
}
