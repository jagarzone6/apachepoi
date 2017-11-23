import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public class POITestNG {

@DataProvider
    Object[][] getXLXSData() throws IOException{

    int columns = 2;
    Object data[][];
    ArrayList<String> labels = new ArrayList<String>();
    ArrayList<String> descriptions = new ArrayList<String>();

    FileInputStream fis = new FileInputStream("src/test/resources/data.xlsx");
    XSSFWorkbook workBook = new XSSFWorkbook(fis);

    XSSFSheet sheet = workBook.getSheet("Sheet1");

    for(int i=1;i<12;i++) {
        XSSFRow row_a = sheet.getRow(i);

        try {
            XSSFCell cell_a = row_a.getCell(1);
            XSSFCell cell_b = row_a.getCell(2);

            //System.out.println(cell_a.getStringCellValue() + " - " + cell_b.getStringCellValue() );
            if(cell_a != null){
                labels.add(cell_a.getStringCellValue());
                descriptions.add(cell_b.getStringCellValue());
            }

        } catch (NullPointerException e) {
            //System.out.println(" ");
        }
    }
    data = new Object[labels.size()][columns];

    for (int i=0;i < labels.size();i++){
        data[i][0]=labels.get(i);
    }
    for (int i=0;i < labels.size();i++){
        data[i][1]=descriptions.get(i);
    }
    System.out.println(data);
    return data;
}

@Test(dataProvider = "getXLXSData")
    public void runTest(String label, String description){

    System.out.println("TEST RUNNED with Excel data: "+ label + " - " + description );


}

}
