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

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

/**
 *
 * @author Mark Bridge <j2eewebtier@gmail.com>
 */
public class ExcelReaderTest {
    
    public ExcelReaderTest() {
    }
    
    @BeforeAll
    public static void setUpClass() {
    }
    
    @AfterAll
    public static void tearDownClass() {
    }
    
    @BeforeEach
    public void setUp() {
    }
    
    @AfterEach
    public void tearDown() {
    }

    @Test
    public void testVisually() throws Exception {
        
        File f = new File(".\\src\\test\\resources\\excel\\excelreadertest.xlsx");
        System.out.println(f.getAbsoluteFile().getPath());
        try {
            ExcelReader reader = new ExcelReader(new File(".\\src\\test\\resources\\excel\\excelreadertest.xlsx"));
            reader.open();
            reader.setSheet("testsheet", true);
            Map<String, Integer> keyMap = reader.getKeyMap();
            
            ArrayList<ArrayList<String>> rowList = reader.getRows();
            
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
