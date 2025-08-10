package dataDriven.excelDataProvider;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class dataProvide {

    DataFormatter formatter = new DataFormatter();

    @Test(dataProvider = "driveTest")
    public void testCaseData(String greeting, String communication, String id){
        System.out.println(greeting);
        System.out.println(communication);
        System.out.println(id);
    }

    @DataProvider(name="driveTest")
    public Object[][] getData() throws IOException {
        //Object[][] data = {{"hello","text","1"},{"bye","message","143"},{"solo","call","453"}};
        //Object data[][] = new Object[2][3];
        //data[0][0]="hello";
        //data[0][1]="text;
        //data[0][2]="1";
        //data[1][0]=bye";
        //data[1][1]="message";
        //data[1][2]="143";

        //tell where the file is located
        FileInputStream fis = new FileInputStream("D:/Storage/Fikry/Course & Bootcamp/QA/Selenium Web Driver Java/excelDataProvider/src/test/java/dataDriven/excelDataProvider/Book 1.xlsx");

        //define workbook intance, because excel considered as a workbook by Apache POI API
        XSSFWorkbook wb = new XSSFWorkbook(fis);

        //ambil sheet
        XSSFSheet sheet = wb.getSheetAt(0);

        //get number of total row present in the sheet
        int rowCount = sheet.getPhysicalNumberOfRows();

        //get first row
        XSSFRow row  = sheet.getRow(0);

        //get last cell number which is the number of coloumn present
        int colCount = row.getLastCellNum();

        //iterate every arrow & throw / create into one array
        Object data[][] = new Object[rowCount-1][colCount];

        for (int i = 0; i < rowCount-1; i++) {
            row = sheet.getRow(i+1); //get data from row
            for (int j = 0; j < colCount; j++) {
                XSSFCell cell = row.getCell(j); //get data from cell
                data[i][j] = formatter.formatCellValue(cell); //convert into string & store it in multidimentional array
            }
        }

        return data;
    }
}
