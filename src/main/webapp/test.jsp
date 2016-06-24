




<%@page import="java.sql.Connection"%>
<%@page import="java.sql.PreparedStatement"%>

<%@page import="org.apache.commons.io.IOUtils"%>
<%@page import="org.apache.commons.*" %>
<%@ page import="org.apache.commons.fileupload.servlet.ServletFileUpload"%>
<%@ page import="org.apache.commons.fileupload.disk.DiskFileItemFactory"%>
<%@ page import="org.apache.commons.fileupload.*"%>


<%@ page import="java.util.*,java.io.*"%>
<%@ page import="java.util.Iterator"%>
<%@ page import="java.util.List"%>
<%@ page import="java.util.Map"%>
<%@ page import="java.io.File"%>
<%@ page import="org.apache.poi.hssf.usermodel.HSSFCell"%>
<%@ page import="org.apache.poi.hssf.usermodel.HSSFRow"%>
<%@ page import="org.apache.poi.hssf.usermodel.HSSFSheet"%>
<%@ page import="org.apache.poi.hssf.usermodel.HSSFWorkbook"%>


<head>
<title>Bulk Upload Page</title>
</head>
<body>

<%
Map<String,String> mp=new HashMap<String,String>();
    try {
        String ImageFile = "";
        String itemName = "";
        boolean isMultipart = ServletFileUpload
                .isMultipartContent(request);
        if (!isMultipart) {
        } else {
            FileItemFactory factory = new DiskFileItemFactory();
            ServletFileUpload upload = new ServletFileUpload(factory);
            List items = null;
            try {
                items = upload.parseRequest(request);
            } catch (FileUploadException e) {
                e.getMessage();
            }

            //Connection conn = new MyDataConnect().giveConnection();
            //Date d = new Date();
            FileItem item = (FileItem) items.get(0);

            //InputStream is=item.getInputStream();
             HSSFWorkbook workbook = new HSSFWorkbook(item.getInputStream());
             String ipaddress=request.getRemoteAddr();
             HSSFSheet sh = (HSSFSheet) workbook.getSheet("Sheet1");
                 Iterator<HSSFRow> rowIterator = sh.rowIterator();
                 %>
                 <table border="4">
                 <%
                 int rowcount=0;
                 System.out.println("xdgdsf");
                 while(rowIterator.hasNext()) {     

                     String user_id="";
                     String name="",email="";
                     String age="";
                    HSSFRow row = rowIterator.next();
                     Iterator<HSSFCell> cellIterator = row.cellIterator();


                     %><tr> <%
                     int colcount=0;
                     while(cellIterator.hasNext()) {
                         HSSFCell cell = cellIterator.next();
                         %><td><%
                         String value="";
                         int no=0;
                         switch(cell.getCellType()) {
                         case HSSFCell.CELL_TYPE_NUMERIC:


                             System.out.print(cell.getNumericCellValue() + "\t\t");
                             if((int) cell.getNumericCellValue()==0){
                                 no=0;
                                 }else{
                                     no=(int) cell.getNumericCellValue();
                                 }
                             String column=mp.get(colcount+"");
                             if(column.trim().equals("user_id")){
                                 out.print(cell.getNumericCellValue() + "\t\t");
                                 user_id=no+"";
                             }else if(column.trim().equals("age")){
                                 out.print(cell.getNumericCellValue() + "\t\t");
                                 age=no+"";
                             }

                             break;
                                 case HSSFCell.CELL_TYPE_STRING:        


                                         System.out.print(cell.getStringCellValue().toString() + "\t\t");
                                         if(cell.getStringCellValue().toString()==null || cell.getStringCellValue().toString()=="" || cell.getStringCellValue().toString().trim().length()==0 ){
                                             value="NA";
                                             }else{
                                                 value=cell.getStringCellValue().toString();
                                             }
                                         if(rowcount!=0){
                                             column=mp.get(colcount+"");
                                           //  System.out.println(mp);
                                            if(column.trim().equals("name")){
                                                out.print(cell.getStringCellValue().toString() + "\t\t");
                                                name=value;
                                            }else if(column.trim().equals("email")){
                                                out.print(cell.getStringCellValue().toString() + "\t\t");
                                                email=value;
                                            }else  if(column.trim().equals("user_id")){
                                             out.print(cell.getNumericCellValue() + "\t\t");
                                             user_id=no+"";
                                         }else if(column.trim().equals("age")){
                                             out.print(cell.getNumericCellValue() + "\t\t");
                                             age=no+"";
                                         }




                                         }
                                         break;       

                         }  
                         if(rowcount==0){
                             mp.put(colcount+"",cell.getStringCellValue().toString().trim());
                         }
                         colcount+=1;
                         %></td><%
                         }
                     System.out.println("\n");
                     rowcount+=1;
                     %></tr><%
                     if(rowcount!=0){
                         String query="insert into bulk_upload(user_id,name,email,age) values('"+user_id+"','"+name+"','"+email+"','"+age+"')";
                      System.out.println(query);

                         PreparedStatement ptmt=conn.prepareStatement(query);

                         int i=ptmt.executeUpdate();

                         if(i==1){
                          System.out.println("Updated ---------        ");
                         }else{
                         System.out.println("Not Updated ---------        "); 
                        }
                         ptmt.close();



                     }
                     conn.close();

                 %></table><%
                 }


        }   


    }catch(Exception e){
        e.printStackTrace();
        System.out.println(e);
    }


%>
</body>
</html>