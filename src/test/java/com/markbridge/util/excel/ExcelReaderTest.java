/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.markbridge.util.excel;

import java.io.File;
import java.util.List;
import java.util.Map;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

/**
 *
 * @author bridgema
 */
public class ExcelReaderTest {
    
    public ExcelReaderTest() {
    }
    
    @BeforeClass
    public static void setUpClass() {
    }
    
    @AfterClass
    public static void tearDownClass() {
    }
    
    @Before
    public void setUp() {
    }
    
    @After
    public void tearDown() {
    }

    @Test
    public void testVisually() throws Exception {
        
        File f = new File(".\\src\\test\\resources\\excel\\excelreadertest.xlsx");
        System.out.println(f.getAbsoluteFile().getPath());
        try {
            ExcelReader reader = new ExcelReader(new File(".\\src\\test\\resources\\excel\\excelreadertest.xlsx"));
            reader.open();
            reader.setSheet("testsheet");
            Map<String, Integer> keyMap = reader.getKeyMap();
            
            List<List<String>> rowList = reader.getRows();
            
            for(List<String> row : rowList) {
                for(String cell : row) {
                    System.out.print(cell + "\t");
                }
                System.out.println("");
            }
            
            
            for(List<String> row : rowList) {
                
                System.out.print(row.get(keyMap.get("String")) + "\t");
                System.out.print(row.get(keyMap.get("General")) + "\t");
                System.out.print(row.get(keyMap.get("SeqNum")) + "\t");
                    
                System.out.println("");
            }
            
            reader.close();
        } catch(Exception ex) {
            System.out.println("Exception: " + ex.getMessage());
        }
    }
    
}
