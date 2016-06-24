
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Rough {
public static void main(String args[]) throws IOException{
    
    FileInputStream fin = new FileInputStream("C:\\test.xlsx");
    XSSFWorkbook wb = new XSSFWorkbook(fin);
    XSSFSheet sheet = wb.getSheetAt(1);
    int RowCountWithNullValue=0, RowCountWithoutNullValue=0;
    for (int i=0;i<1000;i++){
        if (sheet.getRow(i)==null)
            RowCountWithNullValue++;
        else{
            RowCountWithoutNullValue++;
        //  System.out.println(sheet.getRow(0).getCell(0));
        }
    }
    System.out.println(sheet.getLastRowNum());
    System.out.println(RowCountWithNullValue+","+RowCountWithoutNullValue);
  }
}