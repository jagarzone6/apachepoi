import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelDataDriven {

    public static void main(String[] args) throws IOException{


        FileInputStream fis = new FileInputStream("src/test/resources/data.xlsx");
        XSSFWorkbook workBook = new XSSFWorkbook(fis);

        XSSFSheet sheet = workBook.getSheet("Sheet1");

        for(int i=1;i<12;i++) {
            XSSFRow row_a = sheet.getRow(i);

            try {
                XSSFCell cell_a = row_a.getCell(1);
                XSSFCell cell_b = row_a.getCell(2);
                System.out.println(cell_a.getStringCellValue()+" - "+ cell_b.getStringCellValue());
            }catch (NullPointerException e){
                System.out.println(" ");
            }

        }

    }

}


