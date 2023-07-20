package ExcelFile;
// Workbook-->Sheet--Rows-->Cell

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingExcel {
    public static void main(String[] args) {
        XSSFWorkbook workbook= new XSSFWorkbook();
        XSSFSheet sheet= workbook.createSheet("Emp info");
        Object empdata[] []={
                {"EmpID","Name","Job"}
        };
    }
}
