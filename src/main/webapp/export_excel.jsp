<%@page import="org.apache.poi.xssf.usermodel.XSSFCell"%>
<%@page import="org.apache.poi.xssf.usermodel.XSSFRow"%>
<%@page import="org.apache.poi.hssf.usermodel.HSSFDataFormat"%>
<%@page import="org.apache.poi.ss.usermodel.Font"%>
<%@page import="org.apache.poi.ss.usermodel.DataFormat"%>
<%@page import="org.apache.poi.ss.usermodel.CellStyle"%>
<%@page import="java.io.FileOutputStream"%>
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
      Class.forName("com.mysql.jdbc.Driver");
      Connection connect = DriverManager.getConnection("jdbc:mysql://localhost:3306/testtozlu" , "root" , "");
      
      Statement statement = connect.createStatement();
      ResultSet resultSet = statement.executeQuery("select * from test");
      XSSFWorkbook workbook = new XSSFWorkbook(); 
      XSSFSheet spreadsheet = workbook
      .createSheet("n11");
      XSSFRow row=spreadsheet.createRow(1);
      XSSFCell cell;
      cell=row.createCell(1);
      cell.setCellValue("Tarih");
      cell=row.createCell(2);
      cell.setCellValue("Takip_no");
      cell=row.createCell(3);
      cell.setCellValue("firma");
      //cell=row.createCell(4);
      //cell.setCellValue("SALARY");
      //cell=row.createCell(5);
      //cell.setCellValue("DEPT");
      int i=2;
      while(resultSet.next())
      {
         row=spreadsheet.createRow(i);
         cell=row.createCell(1);
         cell.setCellValue(resultSet.getString("Ad"));
         cell=row.createCell(2);
         cell.setCellValue(resultSet.getString("tarih"));
         cell=row.createCell(3);
         cell.setCellValue(resultSet.getString("takip_no"));
         cell=row.createCell(4);
         cell.setCellValue(resultSet.getString("firma_k"));
         //cell=row.createCell(5);
         //cell.setCellValue(resultSet.getString("dept"));
         i++;
      }
      FileOutputStream outt = new FileOutputStream(
      new File("E:\\exceldatabase.xlsx"));
      workbook.write(outt);
      outt.close();
      out.print("exceldatabase.xlsx written successfully");
    
    


    

%>