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
public class ExcelWriterTest {
    
    public ExcelWriterTest() {
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
    public void testManually() throws Exception {
        
        File f = new File(".");
        System.out.println(f.getAbsoluteFile().getPath());
        
        String[] headerList = {"col1", "second", "3rd", "4", "fifth", "sixth", "7th", "8", "col9","tenth"};
        ArrayList<ArrayList<String>> memRows = new ArrayList<>();
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
