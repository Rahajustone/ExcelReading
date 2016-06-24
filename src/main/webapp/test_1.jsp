<%@page import="org.apache.poi.xssf.usermodel.XSSFSheet"%>
<%@page import="java.sql.PreparedStatement"%>
<%@page import="java.sql.DriverManager"%>
<%@page import="java.sql.Connection"%>
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
<%@ page import ="java.sql.*" %>
<%@  page import="javax.sql.*" %>


<% 
         FileInputStream input = new FileInputStream("E:\\test.xlsx");

        XSSFWorkbook wb = new XSSFWorkbook(input);
        XSSFSheet sheet = wb.getSheetAt(0);
        Row row;
         out.print("<table border='1'>");
        
        for(int i=1; i<=sheet.getLastRowNum(); i++){
             out.print("<tr>");
        row = sheet.getRow(i);
        out.println("<td>"+i+"</td>");
       
        String id = String.valueOf(row.getCell(0).getNumericCellValue());
        
        out.println("<td>"+id+"</td>");
        Date date = row.getCell((short)1).getDateCellValue();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        String tarih= sdf.format(date);
         
       out.println("<td>"+tarih+"</td>");
        if(row.getCell(2)!=null){
        String takip_no = row.getCell(2).getStringCellValue();
         out.println("<td>"+takip_no+"</d>");
     
        }
        else{
            String takip_no ="null";
            
             out.println("<td>"+takip_no+"</d>");
            
        }
        
        
        
        String kargo_f = row.getCell(3).getStringCellValue();
         out.println("<td>"+kargo_f+"</td>");
 
        
        
         out.print("</tr>");
        }
       
        out.print("</table>");
  
       

%>

