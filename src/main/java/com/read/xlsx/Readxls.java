package com.read.xlsx;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
public class Readxls {

public static void main(String[] args) throws Exception, FileNotFoundException, IOException {
String filename = "C:\\test.xlsx";

Workbook workbook = WorkbookFactory.create(new FileInputStream(filename)); // or sample.xls
System.out.println("Number Of Sheets" + workbook.getNumberOfSheets());
Sheet sheet = workbook.getSheetAt(0);
System.out.println("Number Of Rows:" + sheet.getLastRowNum());

for (Row row : sheet) {
for (Cell cell : row) {

CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
System.out.print(cellRef.formatAsString());

System.out.print(" – ");

switch (cell.getCellType()) {
case Cell.CELL_TYPE_STRING:
System.out.println(cell.getRichStringCellValue().getString());
break;
case Cell.CELL_TYPE_NUMERIC:
if (DateUtil.isCellDateFormatted(cell)) {
System.out.println(cell.getDateCellValue());
} else {
System.out.println(cell.getNumericCellValue());
}
break;
case Cell.CELL_TYPE_BOOLEAN:
System.out.println(cell.getBooleanCellValue());
break;
case Cell.CELL_TYPE_FORMULA:
System.out.println(cell.getCellFormula());
break;
default:
System.out.println();
}
}
}
/* for (int i=row.getFirstCellNum();i<=row.getLastCellNum();i++){
if( null!=row.getCell(i)){
if(row.getCell(i).getCellType()==0){
System.out.println("Cell Value "+i+" : " + row.getCell(i).getNumericCellValue());
}if(row.getCell(i).getCellType()==1){
System.out.println("Cell Value "+i+" : " + row.getCell(i).getStringCellValue());
System.out.println("Cell Value "+i+" : " + row.getCell(i).getDateCellValue());
}
}
}*/
}

}

