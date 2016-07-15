/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.markbridge.util.excel;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 *
 * @author bridgema
 */
public class Compare {
    
    //TODO complete compare
    public void compare(ArrayList<List<String>> workbook1, int keyColumn1, int column1,
            ArrayList<List<String>> workbook2, int keyColumn2, int column2) {
        
        Map<String, String> keyValues1 = new HashMap<>();
        Map<String, String> keyValues2 = new HashMap<>();
        
        for(List<String> row : workbook1) {
            if(! keyValues1.keySet().contains(row.get(keyColumn1))) {
                keyValues1.put(row.get(keyColumn1), row.get(keyColumn1));
                //...
            } else {
                throw new RuntimeException("key clash: ");
            }
        }
    }
}
