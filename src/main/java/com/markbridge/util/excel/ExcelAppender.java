/*
 * Copyright (c) 2016, Mark Bridge <j2eewebtier@gmail.com>
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 * * Redistributions of source code must retain the above copyright notice, this
 *   list of conditions and the following disclaimer.
 * * Redistributions in binary form must reproduce the above copyright notice,
 *   this list of conditions and the following disclaimer in the documentation
 *   and/or other materials provided with the distribution.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
 * ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
 * LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
 * CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
 * SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
 * INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
 * CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
 * ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
 * POSSIBILITY OF SUCH DAMAGE.
 */
package com.markbridge.util.excel;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
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
 * @author Mark Bridge <j2eewebtier@gmail.com>
 */
public class ExcelAppender {
    
    File file;
    OPCPackage destination;
    Workbook workbook;
    boolean existing = false;
    
    public ExcelAppender(File file) {
        this.file = file;
    }
    
    public ExcelAppender open() throws InvalidFormatException, IOException {
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
     * Clear any existing sheet with given name, create a new one with the given 
     * header at the given position (zero-based) and make it the active sheet
     * 
     * @param sheetName
     * @param sheetPosition 0-based
     * @param header
     * @return 
     */
    public ExcelAppender refreshSheet(String sheetName, 
            Integer sheetPosition, String[] headerArray) throws IOException {
        
        //clear any existing rows on sheet
        Sheet sheet = workbook.getSheet(sheetName);
        if(sheet != null) {
            
            int numSheets = workbook.getNumberOfSheets();
            for(int i = 0; i < numSheets; i++) {
                if(workbook.getSheetAt(i).getSheetName().equalsIgnoreCase(sheetName)) {
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
        workbook.setSheetOrder(sheetName, sheetPosition);
        workbook.setActiveSheet(sheetPosition);
        
        Row row = sheet.createRow(0);
        for(int i = 0; i < headerArray.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(headerArray[i]);
        }
        
        //http://stackoverflow.com/questions/14117617/apache-poi-unable-to-write-to-an-existing-workbook
        BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(file));
        workbook.write(bos);
        bos.flush();
        bos.close();
        
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
    public ExcelAppender appendRows(String sheetName, List<List<String>> memRowList) throws IOException {
        
        Sheet sheet = workbook.getSheet(sheetName);
        
        int lastRowNum = sheet.getLastRowNum();
        
        //add rows in memory - nb, start at row 1 of xssfsheet so plus one on memrownum index 
        //when adding the row to the sheet
        for(int memRowNum = 0; memRowNum < memRowList.size(); memRowNum++) {
            List<String> memRow = memRowList.get(memRowNum);
            Row row = sheet.createRow(memRowNum + lastRowNum + 1);
            for(int i = 0; i < memRow.size(); i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(memRow.get(i));
            }
        }
        
        //http://stackoverflow.com/questions/14117617/apache-poi-unable-to-write-to-an-existing-workbook
        BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(file));
        workbook.write(bos);
        bos.flush();
        bos.close();
        
        return this;
    }
    
    public void close() throws IOException {
        workbook.close();
    }
    
    public Workbook getWorkbook() {
        return workbook;
    }
    
}
