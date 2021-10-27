package com.capgemini.appianrpa.module.excel;

import com.novayre.jidoka.client.api.IJidokaServer;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ModuleExcel {

    private IJidokaServer<?> server;

    public ModuleExcel(){}

    public ModuleExcel(IJidokaServer<?> server) {
        this.server = server;
    }

    public List<HashMap<String, String>> redExcel(String inputFile, String sheetName, int headerIndex, int contentIndex) throws IOException {

        List<HashMap<String, String>> data  = new ArrayList<>();
        try (FileInputStream file = new FileInputStream(inputFile)) {
            //Condition to set file extension
            Workbook wb;
            if (inputFile.endsWith(".xlsx")) {
                wb = new XSSFWorkbook(file);
            } else {
                wb = new HSSFWorkbook(file);
            }
            Sheet sheet = wb.getSheet(sheetName);
            //Iteration to las row with values
            while (sheet.getRow(contentIndex) != null && sheet.getRow(contentIndex).getCell(0) != null && !sheet.getRow(contentIndex).getCell(0).toString().equals("")) {
                Row row = sheet.getRow(contentIndex);
                HashMap<String, String> rowMap = new LinkedHashMap<>();
                int initCell = 0;
                //Iteration to last cell in row with values (header as reference of table)
                while (initCell < sheet.getRow(headerIndex).getLastCellNum()){
                    //Condition to check if cell have no values
                    if(row.getCell(initCell) != null){
                        Cell cell = row.getCell(initCell);
                        rowMap.put(sheet.getRow(headerIndex).getCell(initCell).toString(),getTextFromCell(cell));
                    }else{
                        rowMap.put(sheet.getRow(headerIndex).getCell(initCell).toString(),"");
                    }
                    initCell++;
                }
                data.add(rowMap);
                contentIndex++;
            }
        }
        return data;
    }

    private String getTextFromCell(Cell cell) {
        String value = "";
        //get text by different format cells from Excel
        switch (cell.getCellType()) {
            case STRING:
                value = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    value = cell.getDateCellValue().toString();
                } else {
                    value = String.valueOf(cell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                value = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA:
                value = cell.getCellFormula();
                break;
            case BLANK:
                break;
            default:
        }
        return value;
    }
}
