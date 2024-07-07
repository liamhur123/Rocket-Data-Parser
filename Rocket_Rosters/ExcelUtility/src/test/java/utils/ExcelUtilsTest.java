package utils;
import java.io.IOException;

public class ExcelUtilsTest {
    public static void main(String[] args) throws IOException {

        String PrjDir = System.getProperty("user.dir");
        String excelPath = PrjDir +"/data/active_report.xlsx";
        ExcelUtils excel = new ExcelUtils(excelPath);
        //excel.getAllNames();


        excel.WriteClassNamesandTimes();

        //int rows = excel.getRowCount();
        //System.out.println(rows);
    }

}
