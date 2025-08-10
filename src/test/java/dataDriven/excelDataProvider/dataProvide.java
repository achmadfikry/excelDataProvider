package dataDriven.excelDataProvider;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class dataProvide {

    @Test(dataProvider = "driveTest")
    public void testCaseData(String greeting, String communication, String id){
        System.out.println(greeting);
        System.out.println(communication);
        System.out.println(id);
    }

    @DataProvider(name="driveTest")
    public Object[][] getData() throws IOException {
//        Object[][] data = {{"hello","text","1"},{"bye","message","143"},{"solo","call","453"}};

        //tell where the file is located
        FileInputStream fis = new FileInputStream("/Users/telkomdev/Java/ExcelDriven/src/main/java/Framework/Main.java");

        //define workbook intance, because excel considered as a workbook by Apache POI API
        XSSFWorkbook wb = new XSSFWorkbook(fis);

        //ambil sheet
        XSSFSheet sheet = wb.getSheetAt(0);

        //get number of total row present in the sheet
        int rowCount = sheet.getPhysicalNumberOfRows();

        //get first row
        XSSFRow row  = sheet.getRow(0);

        //get last cell number which is the number of coloumn present
        int colcount = row.getLastCellNum();

        //iterate every arrow & throw / create into one array
        Object data[][] = new Object[rowCount-1][colcount];

        for (int i = 0; i < rowCount; i++) {
            row = sheet.getRow(i+1);
            for (int j = 0; j < colcount; j++) {
                row.getCell(j);
            }
        }

        return data;
    }
}
