package com.github.deansquirrel.tools.poi;

import java.util.List;

/**
 * Excel表格数据
 */
public class WorkTableData {

    private WorkTableData(){};

    private WorkTableData(String name) {
        this.name = name;
    }

    private WorkTableData(String name, List<String> title, List<List<Object>> rows) {
        this.name = name;
        this.title = title;
        this.rows = rows;
    }

    public static WorkTableData builder(String name) {
        return new WorkTableData(name);
    }

    public static WorkTableData builder(String name, List<String> title, List<List<Object>> rows) {
        return new WorkTableData(name, title, rows);
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

    public WorkTableData initTitle(List<String> title) {
        this.title = title;
        return this;
    }

    public List<List<Object>> getRows() {
        return rows;
    }

    public void setRows(List<List<Object>> rows) {
        this.rows = rows;
    }

    public WorkTableData initRows(List<List<Object>> rows) {
        this.rows = rows;
        return this;
    }
}
