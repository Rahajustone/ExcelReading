<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
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
 
         
        Class.forName("com.mysql.jdbc.Driver");//database connection
        Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/testtozlu", "root","");//database connection
        PreparedStatement ps;
        FileInputStream input = new FileInputStream("E:\\test.xlsx");//getting file from a source
        XSSFWorkbook wb = new XSSFWorkbook(input);
        XSSFSheet sheet = wb.getSheetAt(0);
        Row row;
        ps=con.prepareStatement("INSERT INTO test VALUES(?,?,?,?)");//inserting a cell values to a table
       for(int i=1; i<=sheet.getLastRowNum(); i++){
        row = sheet.getRow(i);
        String id = String.valueOf(row.getCell(0).getNumericCellValue());
        ps.setString(1, id);
        Date date = row.getCell((short)1).getDateCellValue();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        String tarih= sdf.format(date);
        ps.setString(2, tarih);
        if(row.getCell(2)!=null){
           String takip_no = row.getCell(2).getStringCellValue();
            ps.setString(3, takip_no);
        }
        else
        {
            String takip_no ="null";
            ps.setString(3, takip_no);
           
        }
        String kargo_f = row.getCell(3).getStringCellValue();
        ps.setString(4, kargo_f);
            

         int cevap=ps.executeUpdate();
        if (cevap > 0) {
           out.println("</br>başari şekilde eklenedi");
                   
        }
         else
        {
             out.println("başarisiz");
        }
               
       
        }
       
   
  
       

%>

