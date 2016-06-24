<%@ page language="java" contentType="text/html; charset=UTF-8"
    pageEncoding="UTF-8"%>
<%response.setContentType("application/vnd.ms-excel");
response.setHeader("Content-disposition", "filename=test.xls" );%>
<%@page import="org.apache.poi.xssf.usermodel.XSSFSheet"%>
<%@page import="java.sql.PreparedStatement"%>
<%@page import="java.sql.DriverManager"%>

<%@ page import ="java.sql.*" %>
<%@  page import="javax.sql.*" %>


<% 
 
         
        Class.forName("com.mysql.jdbc.Driver");
        Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/testtozlu", "root","");
  
        PreparedStatement stmt=con.prepareStatement("select * from test");  
        ResultSet rs=stmt.executeQuery();  
         out.print("<table border='1'>");
         int i=0;
        while(rs.next()){ 
             out.print("<tr>");
               out.print("<td>"+i +"</td>");
                out.print("<td>"+rs.getString(1) +"</td>");
                 out.print("<td>"+rs.getString(2) +"</td>");
                  out.print("<td>"+rs.getString(3) +"</td>");
                 out.print("<td>"+rs.getString(4) +"</td>");
                   out.print("<tr>");
                   i++;
        }  
     out.print("<table>");
         

       
   
  
       

%>

