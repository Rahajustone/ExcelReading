
<%@page import="java.io.FileOutputStream"%>
<%@page import="org.apache.poi.hssf.usermodel.HSSFRichTextString"%>
<%@page import="org.apache.poi.hssf.usermodel.HSSFCell"%>
<%@page import="org.apache.poi.hssf.usermodel.HSSFRow"%>
<%@page import="org.apache.poi.hssf.usermodel.HSSFSheet"%>
<%@page import="java.util.Date"%>
<%@page import="java.text.SimpleDateFormat"%>
<%@page import="org.apache.poi.ss.usermodel.DateUtil"%>
<%@page import="org.apache.poi.hssf.util.CellReference"%>
<%@page import="org.apache.poi.hssf.usermodel.HSSFWorkbook"%>
<%@page import="org.apache.poi.poifs.filesystem.NPOIFSFileSystem"%>
<%@page import="org.apache.poi.ss.usermodel.Cell"%>
<%@page import="java.util.Iterator"%>
<%@page import="java.util.Iterator"%>
<%@page import="org.apache.poi.ss.usermodel.Row"%>
<%@page import="org.apache.poi.ss.usermodel.Sheet"%>
<%@page import="org.apache.poi.xssf.usermodel.XSSFWorkbook"%>
<%@page import="org.apache.poi.ss.usermodel.Workbook"%>
<%@page import="java.io.FileInputStream"%>
<%@page import="java.io.File"%>
<%@page import="org.apache.poi.*" %>


<% 

 

     
    
        //String excelFilePath = "C:\\test.xlsx";
         //Workbook wb = new HSSFWorkbook();
        //FileOutputStream fileOut = new FileOutputStream("C:\\test.xlsx");
        NPOIFSFileSystem fs = new NPOIFSFileSystem(new File("C:\\test.xlsx"));
        HSSFWorkbook wb = new HSSFWorkbook(fs.getRoot(), true);
         Sheet sheet1 = wb.getSheetAt(0);
    for (Row row : sheet1) {
        for (Cell cell : row) {
            CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
            System.out.print(cellRef.formatAsString());
            System.out.print(" - ");

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

%>);

