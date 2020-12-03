/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.camilledesmots.androidcalllogstats;

import java.io.IOException;
import static java.lang.System.exit;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.camilledesmots.CSVtoObject;

/**
 *
 * @author camilledesmots
 */
public class AndroidCallLogStats {

    private static final Logger LOG = Logger.getLogger(AndroidCallLogStats.class.getName());

    public static void main(String args[]) {

        System.out.println("AndroidCallLogStats");
        System.out.println("===================");
        System.out.println("Création de statistiques à partir des sauvegardes d'historiques");
        System.out.println("à Androïd à partir des sauvegardes au format .clbu");
        System.out.println("");
        

        String clbuFolder = System.getProperty("clbuFolder", "");
        String resultFolder = System.getProperty("resultFolder", "");

        if (clbuFolder.isEmpty() && resultFolder.isEmpty()) {
            System.out.println("Appeler le programme avec les options -DclubFolder et -DresultFolder");
            System.out.println(" -DclbuFolder=   est les répertoire ou sont stockés les fichiers avec l'extension .clbu");
            System.out.println(" -DresultFolder= est le répertoire ou sera généré le fichier EXCEL");
            System.out.println("");
            System.out.println("Exemple :");
            System.out.println("java -DclbuFoler=\"C:\\tmp\" -DresultFolder=\"C:\\Mon répertoire\"");
            
            exit(-1);

        }

        System.out.println("-DclbuFolder=");
        System.out.println("   " + clbuFolder);
        System.out.println("-DresultFolder=");
        System.out.println("   " + resultFolder);
        System.out.println("");
        
        CSVtoObject csvToObject = new CSVtoObject();

        try {
            csvToObject.readDirectory(clbuFolder, resultFolder);
            //TODO Continuer ici
        } catch (IOException ex) {
            Logger.getLogger(AndroidCallLogStats.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

}
