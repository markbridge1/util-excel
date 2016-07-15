/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.markbridge.util.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author bridgema
 */
public class ExcelWriter {
    
    File file;
    OPCPackage destination;
    Workbook workbook;
    boolean existing = false;
    
    public ExcelWriter(File file) {
        this.file = file;
    }
    
    public ExcelWriter open() throws InvalidFormatException, IOException {
        if(file.exists()) {
            existing = Boolean.TRUE;
            FileInputStream fos = new FileInputStream(file);
            destination = OPCPackage.open(fos);
            workbook = new XSSFWorkbook(destination);
            fos.close();
        } else {
            workbook = new XSSFWorkbook();
        }
        return this;
    }
    
    /**
     * https://poi.apache.org/spreadsheet/quick-guide.html
     * 
     * Assumes have a header row and at least one row of data, will write and close everything out
     * @param sheetName
     * @param headersList
     * @param memRowList 
     */
    public void writeRows(String sheetName, 
            String[] headersList, ArrayList<List<String>> memRowList) throws IOException {
        
        int sheetOrder = 0; //default to first unless already have a sheet
        
        //clear any existing rows
        Sheet sheet = workbook.getSheet(sheetName);
        if(sheet != null) {
            int numSheets = workbook.getNumberOfSheets();
            for(int i = 0; i < numSheets; i++) {
                if(workbook.getSheetAt(i).getSheetName().equalsIgnoreCase(sheetName)) {
                    sheetOrder = i;
                    sheet = workbook.getSheetAt(i);
                    for(int rowNum = sheet.getFirstRowNum(); rowNum <= sheet.getLastRowNum(); rowNum++ ) {
                        sheet.removeRow(sheet.getRow(rowNum));
                    }
                    workbook.removeSheetAt(i);
                    break;
                }
            }
        }
        
        sheet = workbook.createSheet(sheetName);
        workbook.setSheetOrder(sheetName, sheetOrder);
        workbook.setActiveSheet(sheetOrder);
        
        //add header row
        Row headerRow = sheet.createRow(0);
        for(int i = 0; i < headersList.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue(headersList[i]);
        }
        
        //add rows in memory - nb, start at row 1 of xssfsheet so plus one on memrownum index 
        //when adding the row to the sheet
        for(int memRowNum = 0; memRowNum < memRowList.size(); memRowNum++) {
            List<String> memRow = memRowList.get(memRowNum);
            Row row = sheet.createRow(memRowNum + 1);
            for(int i = 0; i < memRow.size(); i++) {
                Cell cell = row.createCell(i);
                cell.setCellType(Cell.CELL_TYPE_STRING);
                cell.setCellValue(memRow.get(i));
            }
        }
        
        //http://stackoverflow.com/questions/14117617/apache-poi-unable-to-write-to-an-existing-workbook
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos);
        fos.flush();
        fos.close();
    }
    
    public void close() throws IOException {
        workbook.close();
    }
    
}
