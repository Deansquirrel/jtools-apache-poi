package com.github.deansquirrel.tools.poi;

import java.util.List;

public interface IDataMapper<T> {

    List<String> getTitleList();
    List<Object> getRowData(T data);

}
