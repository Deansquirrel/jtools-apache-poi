package com.github.deansquirrel.tools.poi;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;

public class WorkBookTool {

    private static final Logger logger = LoggerFactory.getLogger(WorkBookTool.class);

    private WorkBookTool(){}

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
     * 获取单元格样式
     * @param book 表格对象
     * @param font 字体
     * @param format 格式
     * @return 单元格样式
     */
    protected static CellStyle getCellStyle(Workbook book, Font font, String format) {
        if(book == null) {
            return null;
        }
        CellStyle cellStyle = book.createCellStyle();
        cellStyle.setFont(font == null ? getFont(book) : font);
        if(format != null && !format.isEmpty()) {
            CreationHelper creationHelper = book.getCreationHelper();
            cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat(format));
        }
        return cellStyle;
    }

    /**
     * 获取单元格样式
     * @param book 表格对象
     * @param font 字体
     * @return 单元格样式
     */
    protected static CellStyle getCellStyle(Workbook book, Font font) {
        return getCellStyle(book, font, null);
    }

    /**
     * 获取单元格样式
     * @param book 表格对象
     * @return 单元格样式
     */
    protected static CellStyle getCellStyle(Workbook book) {
        return getCellStyle(book, null, null);
    }

    /**
     * 获取数据格式字符串
     * @param l 小数保留位数
     * @return 格式字符串
     */
    public static String getNumberFormat(int l) {
        if(l <= 0) {
            return "##0";
        }
        StringBuilder sb = new StringBuilder();
        sb.append("###0").append(".");
        for(int i = 0; i < l; i++) {
            sb.append("0");
        }
        return sb.toString();
    }

    /**
     * 获取百分比字符串格式
     * @return 百分比字符串格式
     */
    public static String getPercentFormat() {
        return "0.00%";
    }

    private static Map<Integer, Map<Integer, CellStyle>> getCellStyleMap(
            Workbook book, Font font,
            Map<Integer, Map<Integer, String>> map) {
        Map<Integer, Map<Integer, CellStyle>> resultCellStyleMap = new HashMap<>();
        if(map == null || map.isEmpty()) {
            return resultCellStyleMap;
        }
        for(Integer key : map.keySet()) {
            Map<Integer, String> subMap = map.get(key);
            if(subMap == null || subMap.isEmpty()) {
                continue;
            }
            Map<Integer, CellStyle> subCellStyleMap = new HashMap<>();
            for(Integer subKey : subMap.keySet()) {
                String format = subMap.get(subKey);
                if(format == null || format.isEmpty()) {
                    continue;
                }
                CellStyle cellStyle = getCellStyle(book, font, format);
                if(cellStyle == null) {
                    continue;
                }
                subCellStyleMap.put(subKey, cellStyle);
            }
            resultCellStyleMap.put(key, subCellStyleMap);
        }
        return resultCellStyleMap;
    }

    private static final String DATE_ZERO = "000000";
    private static final SimpleDateFormat checkHMS = new SimpleDateFormat("HHmmss");

    /**
     * 数据数据生成文件
     * @param list 数据列表
     * @return 文件对象
     */
    public static SXSSFWorkbook getSXSSFWorkBook(List<WorkTableData> list) {
        return getSXSSFWorkBook(list, null);
    }

    /**
     * 数据数据生成文件
     * @param list 数据列表
     * @param dataFormat 特定格式，对应数据起始编号为零
     * @return 文件对象
     */
    public static SXSSFWorkbook getSXSSFWorkBook(List<WorkTableData> list, Map<Integer, Map<Integer, String>> dataFormat) {
        SXSSFWorkbook book = new SXSSFWorkbook();
        if(list == null || list.isEmpty()) {
            book.createSheet();
        } else {
            Font font = getFont(book);
            CellStyle defaultCellStyle = getCellStyle(book, font);
            CellStyle numberCellStyle_0 = getCellStyle(book, font, getNumberFormat(0));
            CellStyle numberCellStyle_2 = getCellStyle(book, font, getNumberFormat(2));
            CellStyle dateCellStyle = getCellStyle(book, font, "yyyy-MM-dd");
            CellStyle dateTimeCellStyle = getCellStyle(book, font, "yyyy-mm-dd hh:mm:ss");
            Map<Integer, Map<Integer, CellStyle>> cellStyleMap = getCellStyleMap(book, font, dataFormat);
            for(int tableIndex = 0; tableIndex < list.size(); tableIndex++) {
                WorkTableData table = list.get(tableIndex);
                if(table == null) {
                    continue;
                }
                SXSSFSheet sheet = book.createSheet(getNextSheetName(book, table.getName()));
                List<String> title = table.getTitle();
                if(title != null && !title.isEmpty()) {
                    Row row = sheet.createRow(0);
                    for(int i = 0; i < title.size(); i++) {
                        Cell cell = row.createCell(i);
                        cell.setCellValue(new XSSFRichTextString(title.get(i)));
                        cell.setCellStyle(defaultCellStyle);
                    }
                }
                List<List<Object>> dataList = table.getRows();
                if(dataList != null && !dataList.isEmpty()) {
                    Map<Integer, CellStyle> cellStyleSubMap = cellStyleMap.getOrDefault(tableIndex, new HashMap<>());
                    int rowIndex = 1;
                    for(List<Object> rowData : dataList) {
                        if(rowData == null || rowData.isEmpty()) {
                            //跳过空数据行（空数据行不增加行号）
                            continue;
                        }
                        Row row = sheet.createRow(rowIndex++);
                        for(int i = 0; i < rowData.size(); i++) {
                            Object d = rowData.get(i);
                            Cell cell = row.createCell(i);
                            CellStyle cellStyle = cellStyleSubMap.getOrDefault(i, null);
                            if(d == null) {
                                continue;
                            }
                            if(d instanceof String) {
                                cell.setCellValue(new XSSFRichTextString(String.valueOf(d)));
                                cell.setCellStyle(defaultCellStyle);
                            } else if(d instanceof BigDecimal) {
                                cell.setCellValue(Double.parseDouble(((BigDecimal) d).toPlainString()));
                                cell.setCellStyle(cellStyle == null ? numberCellStyle_2 : cellStyle);
                            } else if(d instanceof Integer || d instanceof Long) {
                                cell.setCellValue(Long.parseLong(String.valueOf(d)));
                                cell.setCellStyle(cellStyle == null ? numberCellStyle_0 : cellStyle);
                            } else if(d instanceof Number) {
                                cell.setCellValue(Double.parseDouble(String.valueOf(d)));
                                cell.setCellStyle(cellStyle == null ? numberCellStyle_2 : cellStyle);
                            } else if(d instanceof Date) {
                                cell.setCellValue((Date) d);
                                if (cellStyle == null) {
                                    if (DATE_ZERO.equals(checkHMS.format(d))) {
                                        cell.setCellStyle(dateCellStyle);
                                    } else {
                                        cell.setCellStyle(dateTimeCellStyle);
                                    }
                                } else {
                                    cell.setCellStyle(cellStyle);
                                }
                            } else {
                                cell.setCellValue(String.valueOf(d));
                                cell.setCellStyle(cellStyle == null ? defaultCellStyle : cellStyle);
                            }
                        }
                    }
                }
            }
        }
        return book;
    }

    protected static String getNextSheetName(SXSSFWorkbook book, String tableName) {
        String tn = tableName == null ? "" : tableName.trim();
        if(!tn.isEmpty()) {
            return tn;
        }
        if(book == null) {
            return String.valueOf(System.currentTimeMillis());
        }
        String sheetName = "Sheet" + Math.max(book.getNumberOfSheets(), 1);
        for(int idx = 1; book.getSheet(sheetName) != null; ++idx) {
            sheetName = "Sheet" + idx;
        }
        return sheetName;
    }

    private static final long MAX_LINE_SIZE = (long) 100 * 10000;

    /**
     * 获取单页数据对象
     * @param name sheet页名称
     * @param list 数据列表
     * @param iDataMapper 数据展示配置
     * @return 单页数据对象
     * @param <T> 数据类型
     */
    public static <T> WorkTableData getXSSFWorkTable(String name, List<T> list, IDataMapper<T> iDataMapper) {
        List<List<Object>> rows = new ArrayList<>();
        if(list != null) {
            if(list.size() > MAX_LINE_SIZE) {
                throw new ArrayIndexOutOfBoundsException("数据长度超长 " + list.size());
            }
            for(T data : list) {
                rows.add(iDataMapper.getRowData(data));
            }
        }
        return WorkTableData.builder(name)
                .initTitle(iDataMapper.getTitleList())
                .initRows(rows);
    }

    /**
     * 获取单页数据对象
     * @param name sheet页名称
     * @param data 数据列表（纯字符串类型）
     * @return 单页数据对象
     */
    public static WorkTableData getWorkTableData(String name, String[][] data) {
        if(data == null || data.length == 0) {
            return null;
        }
        if(data.length > MAX_LINE_SIZE) {
            throw new ArrayIndexOutOfBoundsException("数据长度超长 " + data.length);
        }
        String[] title = data[0];
        List<List<Object>> rows = new ArrayList<>();
        for(int i = 1; i < data.length; i++) {
            String[] row = data[i];
            List<Object> rowList = new ArrayList<>(Arrays.asList(row));
            rows.add(rowList);
        }
        return WorkTableData.builder(name, Arrays.asList(title), rows);
    }

    /**
     * 获取单页数据对象
     * @param rows 数据列表（纯字符串类型）
     * @return 单页数据对象
     */
    public static WorkTableData transWorkTableData(String[][] rows) {
        return getWorkTableData(null, rows);
    }

    /**
     * 保存文件对象
     * @param base 基础路径
     * @param fileName 文件名称
     * @param f 文件对象
     * @throws IOException 异常
     */
    public static void saveSXSSFWorkbook(String base, String fileName, SXSSFWorkbook f) throws IOException {
        String fullPath = base + ((base == null || base.isEmpty()) ? "" : File.separator) + fileName;
        try (FileOutputStream fs = new FileOutputStream(fullPath)) {
            f.write(fs);
            fs.flush();
        }
    }

    /**
     * 获取文件名称
     * @param nameList 内容
     * @param separator 分隔符
     * @return 文件名称
     */
    public static String getXLSXFileName(List<String> nameList, String separator) {
        if(nameList == null || nameList.isEmpty()) {
            return System.currentTimeMillis() + ".xlsx";
        }
        List<String> list = new ArrayList<>();
        for(String name : nameList) {
            if(name == null || name.isEmpty() || name.trim().isEmpty()) {
                continue;
            }
            list.add(name.trim());
        }
        String sep = (separator == null || separator.isEmpty() || separator.trim().isEmpty()) ? "-" : separator.trim();
        return StringUtils.join(list, sep)  + ".xlsx";
    }

    /**
     * 获取文件名称
     * @param nameList 内容
     * @return 文件名称
     */
    public static String getXLSXFileName(List<String> nameList) {
        return getXLSXFileName(nameList, "-");
    }

}
