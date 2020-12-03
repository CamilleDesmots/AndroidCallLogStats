/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.camilledesmots;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Camille Desmots
 */
public class TestXSS {

    // Create a Logger 
    private static final Logger LOG = Logger.getLogger(CSVtoObject.class.getName());

    public static void main(String[] args) {
        //or new HSSFWorkbook();
        Workbook wb = new XSSFWorkbook();
        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet("new sheet");

        // Create a row and put some cells in it. ExcelRows are 0 based.
        Row row = sheet.createRow((short) 0);
        // Create a cell and put a value in it.
        Cell cell = row.createCell((short) 0);
        cell.setCellValue(1);

        //numeric value
        row.createCell(1).setCellValue(1.2);

        //plain string value
        row.createCell(2).setCellValue("This is a string cell");

        //rich text string
        RichTextString str = creationHelper.createRichTextString("Apache");
        Font font = wb.createFont();
        font.setItalic(true);
        font.setUnderline(Font.U_SINGLE);
        str.applyFont(font);
        row.createCell(3).setCellValue(str);

        //boolean value
        row.createCell(4).setCellValue(true);

        //formula
        row.createCell(5).setCellFormula("SUM(A1:B1)");

        //date
        CellStyle style = wb.createCellStyle();
        style.setDataFormat(creationHelper.createDataFormat().getFormat("m/d/yy h:mm"));
        cell = row.createCell(6);
        cell.setCellValue(new Date());
        cell.setCellStyle(style);

        //hyperlink
        row.createCell(7).setCellFormula("SUM(A1:B1)");
        cell.setCellFormula("HYPERLINK(\"http://google.com\",\"Google\")");

        // Write the output to a file
        try ( FileOutputStream fileOut = new FileOutputStream("target/ooxml-cell.xlsx")) {
            wb.write(fileOut);
        } catch (IOException ex) {
            LOG.log(Level.SEVERE, null, ex);

        }
    }
}