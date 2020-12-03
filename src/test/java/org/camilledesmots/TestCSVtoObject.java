/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.camilledesmots;

import java.io.IOException;

/**
 *
 * @author root
 */
public class TestCSVtoObject {
     public static void main(String args[]) throws IOException {
         
         CSVtoObject test;
         test = new CSVtoObject();
         
         test.readDirectory("/Users/camilledesmots/Google Drive/oed.lpo35/Historique des appels/Log des appels/", "target");
     }
    
}
