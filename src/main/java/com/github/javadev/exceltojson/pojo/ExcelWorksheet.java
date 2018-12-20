package com.github.javadev.exceltojson.pojo;

import java.util.ArrayList;
import java.util.List;

public class ExcelWorksheet {

    private String name;
    private List<List<Object>> data = new ArrayList<List<Object>>();
    private int maxCols = 0;

    public void addRow(ArrayList<Object> row) {
        data.add(row);
        if (maxCols < row.size()) {
            maxCols = row.size();
        }
    }

    public int getMaxRows() {
        return data.size();
    }

    public void fillColumns() {
        for (List<Object> tmp : data) {
            while (tmp.size() < maxCols) {
                tmp.add(null);
            }
        }
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<List<Object>> getData() {
        return data;
    }

    public void setData(List<List<Object>> data) {
        this.data = data;
    }

    public int getMaxCols() {
        return maxCols;
    }

    public void setMaxCols(int maxCols) {
        this.maxCols = maxCols;
    }
}
