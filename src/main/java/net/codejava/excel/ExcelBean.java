/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package net.codejava.excel;
import net.codejava.excel.SimpleExcelReaderExample;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author rahaa
 */
public class ExcelBean {
   
 
    private double title;
    private String tarih;
    private float takip_no;
    private String kargafirma;
 
    public ExcelBean() {
    }
 
    public String toString() {
        return String.format("%f - %s - %f - %s ", title, tarih, takip_no,kargafirma);
    }
 
    // getters and setters
    private Object getCellValue(Cell cell) {
    switch (cell.getCellType()) {
    case Cell.CELL_TYPE_STRING:
        return cell.getStringCellValue();
 
    case Cell.CELL_TYPE_BOOLEAN:
        return cell.getBooleanCellValue();
 
    case Cell.CELL_TYPE_NUMERIC:
        return cell.getNumericCellValue();
    }
 
    return null;
}

    
}
