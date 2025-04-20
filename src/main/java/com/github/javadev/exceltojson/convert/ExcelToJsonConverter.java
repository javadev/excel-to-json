package com.github.javadev.exceltojson.convert;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;

import com.github.javadev.exceltojson.pojo.ExcelWorkbook;
import com.github.javadev.exceltojson.pojo.ExcelWorksheet;

import org.apache.poi.ss.usermodel.DateUtil;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelToJsonConverter {

    private ExcelToJsonConverterConfig config = null;

    public ExcelToJsonConverter(ExcelToJsonConverterConfig config) {
        this.config = config;
    }

    public static ExcelWorkbook convert(ExcelToJsonConverterConfig config) throws IOException {
        return new ExcelToJsonConverter(config).convert();
    }

    public ExcelWorkbook convert()
            throws IOException {
        ExcelWorkbook book = new ExcelWorkbook();

        InputStream inp = Files.newInputStream(Paths.get(config.getSourceFile()));
        Workbook wb = WorkbookFactory.create(inp);

        book.setFileName(config.getSourceFile());

        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            if (sheet == null) {
                continue;
            }
            ExcelWorksheet tmp = new ExcelWorksheet();
            tmp.setName(sheet.getSheetName());
            for (int j = sheet.getFirstRowNum(); j <= sheet.getLastRowNum(); j++) {
                Row row = sheet.getRow(j);
                if (row == null) {
                    continue;
                }
                boolean hasValues = false;
                ArrayList<Object> rowData = new ArrayList<Object>();
                for (int k = 0; k <= row.getLastCellNum(); k++) {
                    Cell cell = row.getCell(k);
                    if (cell != null) {
                        Object value = cellToObject(cell);
                        hasValues = hasValues || value != null;
                        rowData.add(value);
                    } else {
                        rowData.add(null);
                    }
                }
                if (hasValues || !config.isOmitEmpty()) {
                    tmp.addRow(rowData);
                }
            }
            if (config.isFillColumns()) {
                tmp.fillColumns();
            }
            book.addExcelWorksheet(tmp);
        }

        return book;
    }

    private Object cellToObject(Cell cell) {

        CellType type = cell.getCellType();

        if (type == CellType.STRING) {
            return cleanString(cell.getStringCellValue());
        }

        if (type == CellType.BOOLEAN) {
            return cell.getBooleanCellValue();
        }

        if (type == CellType.NUMERIC) {

            if (cell.getCellStyle().getDataFormatString().contains("%")) {
                return cell.getNumericCellValue() * 100;
            }

            return numeric(cell);
        }

        if (type == CellType.FORMULA) {
            switch (cell.getCachedFormulaResultType()) {
                case NUMERIC:
                    return numeric(cell);
                case STRING:
                    return cleanString(cell.getRichStringCellValue().toString());
            }
        }

        return null;

    }

    private String cleanString(String str) {
        return str.replace("\n", "").replace("\r", "");
    }

    private Object numeric(Cell cell) {
        if (DateUtil.isCellDateFormatted(cell)) {
            if (config.getFormatDate() != null) {
                return config.getFormatDate().format(cell.getDateCellValue());
            }
            return cell.getDateCellValue();
        }
        return cell.getNumericCellValue();
    }
}
