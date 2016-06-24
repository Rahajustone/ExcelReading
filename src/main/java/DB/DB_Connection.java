 package DB;  
 import java.sql.Connection;  
 import java.sql.DriverManager;  
 import java.sql.SQLException;  
 public class DB_Connection {  
   private Connection con;  
    public DB_Connection()  
   {  
          try  
         {  
               String conUrl="jdbc:mysql://localhost:3306/stud";
               String userName="root";  
               String pass="";  
               Class.forName("com.mysql.jdbc.Driver");  
                 con=DriverManager.getConnection(conUrl,userName,pass);  
         }  
         catch(SQLException s)  
         {  
             System.out.println(s);  
         }  
         catch(ClassNotFoundException c)  
         {  
              System.out.println(c);  
         }  
   }  
   public Connection getConn() {  
     return con;  
   }  
   public void setConn(Connection con) {  
     this.con = con;  
   }  
 }  