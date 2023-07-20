package ExcelFile;





import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;


public class Students {
    public static void main(String[] args) {



    public static void main(String[] args) throws IOException {
        File excelFilePath= new File("C:\\Users\\viswa\\IntelliJ_work_space\\Class_Work_18_0_2023\\src\\main\\resources\\students.xlsx");
        FileInputStream inputStream=new FileInputStream(excelFilePath);
        XSSFWorkbook xssfWorkbook= new XSSFWorkbook(inputStream);
        XSSFSheet sheet=xssfWorkbook.getSheetAt(0);

        int rows= sheet.getLastRowNum();
        int cols=sheet.getRow(1).getLastCellNum();

        for(int r=0;r<=rows;r++)
        {
            XSSFRow row =sheet.getRow(r);
            for(int c=0;c<cols;c++)
            {
                XSSFCell cell=row.getCell(c);

                switch (cell.getCellType()) //type of the cell could be string,boolean or any type
                {
                    case STRING : System.out.print(cell.getStringCellValue()); break;
                    case NUMERIC:System.out.print(cell.getNumericCellValue()); break;
                    case BOOLEAN:System.out.print(cell.getBooleanCellValue()); break;
                }
                System.out.print(" | ");
            }
            System.out.println( );


            for (int j = 0; j < 2; j++) {
                Row row = xssfSheet.createRow(xssfSheet.getLastRowNum() + 1);
                for (int i = 0; i < 3; i++) {
                    Cell cell = row.createCell(i);
                    switch (i) {
                        case 0:
                            if (j == 0) {
                                cell.setCellValue("F");
                            } else if (j == 1) {
                                cell.setCellValue("G");
                            }
                            break;
                        case 1:
                            if (j == 0) {
                                cell.setCellValue(15);
                            } else if (j == 1) {
                                cell.setCellValue(14);
                            }
                            break;
                        case 2:
                            if (j == 0) {
                                cell.setCellValue(5);
                            } else if (j == 1) {
                                cell.setCellValue(7);
                            }
                            break;
                    }
                }
            }

            //write back to excel
            FileOutputStream fileOutputStream = new FileOutputStream(file);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            System.out.println("Successfully Saved to the excel file!!");

        } catch (Exception e) {
                System.err.println("Error is :" +e.getMessage());
            }
        }
    }




    }

    }

