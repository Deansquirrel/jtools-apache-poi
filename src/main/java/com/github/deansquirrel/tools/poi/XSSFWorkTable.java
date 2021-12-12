package com.github.deansquirrel.tools.poi;

import java.util.List;

/**
 * Excel表格数据
 */
public class XSSFWorkTable {

    private XSSFWorkTable(){};

    private XSSFWorkTable(String name) {
        this.name = name;
    }

    private XSSFWorkTable(String name, List<String> title, List<List<Object>> rows) {
        this.name = name;
        this.title = title;
        this.rows = rows;
    }

    public static XSSFWorkTable builder(String name) {
        return new XSSFWorkTable(name);
    }

    public static XSSFWorkTable builder(String name, List<String> title, List<List<Object>> rows) {
        return new XSSFWorkTable(name, title, rows);
    }

    private String name;
    private List<String> title;
    private List<List<Object>> rows;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<String> getTitle() {
        return title;
    }

    public void setTitle(List<String> title) {
        this.title = title;
    }

    public XSSFWorkTable initTitle(List<String> title) {
        this.title = title;
        return this;
    }

    public List<List<Object>> getRows() {
        return rows;
    }

    public void setRows(List<List<Object>> rows) {
        this.rows = rows;
    }

    public XSSFWorkTable initRows(List<List<Object>> rows) {
        this.rows = rows;
        return this;
    }
}
