/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.markbridge.util.excel;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author bridgema
 */
public class ExcelReader {
    
    File file;
    OPCPackage source;
    XSSFWorkbook workbook;
    FormulaEvaluator formulaEvaluator;
    DataFormatter dataFormatter = new DataFormatter(true);
    XSSFSheet sheet;
    List<XSSFRow> rowList;
    Map<String, Integer> keyMap;
    
    public ExcelReader(File file) {
        this.file = file;
    }
    
    public ExcelReader open() throws InvalidFormatException, IOException {
        source = OPCPackage.open(file);
        workbook = new XSSFWorkbook(source);
        formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
        formulaEvaluator.evaluateAll();
        return this;
    }
    
    public ExcelReader setSheet(String sheetname) {
        sheet = workbook.getSheet(sheetname);
        setRowList();
        setKeyMap();
        return this;
    }
    
    public XSSFSheet getSheet() {
        return sheet;
    }
    
    public ArrayList<ArrayList<String>> getRows() {
        ArrayList<ArrayList<String>> retVal = new ArrayList<ArrayList<String>>();
        
        for(XSSFRow row : rowList) {
            ArrayList<String> cellList = new ArrayList<>();
            retVal.add(cellList);
            for(int cellNum = row.getFirstCellNum(); cellNum <= row.getLastCellNum(); cellNum++) {
                cellList.add(asString(row.getCell(cellNum)));
            }
        }
        
        return retVal;
    }
    
    public Map<String, Integer> getKeyMap() {
        return keyMap;
    }
    
    public void close() throws IOException {
        workbook.close();
    }
    
    /**
     * Read in the rows in the set sheet
     * @return list of rows in the sheet set in this instance
     */
    private List<XSSFRow> setRowList() {
        rowList = new ArrayList<XSSFRow>();
        for(int rowNum = sheet.getFirstRowNum(); rowNum <= sheet.getLastRowNum(); rowNum++) {
            rowList.add(sheet.getRow(rowNum));
        }
        return rowList;
    }
    
    //http://stackoverflow.com/questions/1072561/how-can-i-read-numeric-strings-in-excel-cells-as-string-not-numbers-with-apach
    private String asString(XSSFCell cell) {
        return dataFormatter.formatCellValue(cell, formulaEvaluator);
    }
    
    /**
     * Fields (column names) are assumed to be in the first row and unique in the sheet row
     * @return 
     */
    private Map<String, Integer> setKeyMap() {
        
        keyMap = new HashMap<>();
        
        XSSFRow row = sheet.getRow(sheet.getFirstRowNum());
        
        for(int cellNum = 0; cellNum <= row.getLastCellNum(); cellNum++) {
            if(row.getCell(cellNum) == null) {
                continue;
            }
            Integer put = keyMap.put(row.getCell(cellNum).getStringCellValue().trim(), cellNum);
            if(put != null) {
                throw new RuntimeException("Sheet Name: " + sheet.getSheetName()
                        + " Duplicate key: " + row.getCell(cellNum).getStringCellValue().trim());
            }
        }
        
        return keyMap;
    }
}
