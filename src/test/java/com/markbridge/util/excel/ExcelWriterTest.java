/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.markbridge.util.excel;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

/**
 *
 * @author bridgema
 */
public class ExcelWriterTest {
    
    public ExcelWriterTest() {
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
    public void testManually() throws Exception {
        
        File f = new File(".");
        System.out.println(f.getAbsoluteFile().getPath());
        
        String[] headerList = {"col1", "second", "3rd", "4", "fifth", "sixth", "7th", "8", "col9","tenth"};
        ArrayList<List<String>> memRows = new ArrayList<>();
        for(int rowNum = 0; rowNum < 50; rowNum++) {
            ArrayList<String> row = new ArrayList<String>();
            memRows.add(row);
            for(int cellNum = 0; cellNum < 10; cellNum++) {
                row.add("r: " + rowNum + " cell:" + cellNum);
            }
        }
        
        try {
            ExcelWriter writer = new ExcelWriter(new File(".\\src\\test\\resources\\excel\\excelreadertest.xlsx"));
            writer.open();
            writer.writeRows("testsheet", headerList, memRows);
            writer.writeRows("testsheet2", headerList, memRows);
            writer.close();
        } catch(Exception ex) {
            System.out.println("Exception: " + ex.getMessage());
            throw ex;
        }
    }
    
}
