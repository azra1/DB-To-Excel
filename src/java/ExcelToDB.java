/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author Azra
 */

    
import java.io.File;
import java.io.FileOutputStream;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelToDB {
   public static void main(String[] args) throws Exception {
       String myFile="excel.xlsx";
      Class.forName("com.mysql.jdbc.Driver");
      Connection connect = DriverManager.getConnection( 
         "jdbc:mysql://localhost:3306/db1" , 
         "root" , 
         "root"
      );
      
      Statement statement = connect.createStatement();
      ResultSet resultSet = statement.executeQuery("select * from excelTable");
      XSSFWorkbook workbook = new XSSFWorkbook(); 
      XSSFSheet spreadsheet = workbook.createSheet("excel db");
      
      XSSFRow row = spreadsheet.createRow(1);
      XSSFCell cell;
      cell = row.createCell(1);
      cell.setCellValue("ANI");
      cell = row.createCell(2);
      cell.setCellValue("Call Date Time");
      cell = row.createCell(3);
      cell.setCellValue("Duration");
      cell = row.createCell(4);
      cell.setCellValue("Id");
      cell = row.createCell(5);
      cell.setCellValue("Shortcode");
      cell=row.createCell(6);
      cell.setCellValue("Portal");
      cell=row.createCell(7);
      cell.setCellValue("Click Id");
      int i = 2;

      while(resultSet.next()) {
         row = spreadsheet.createRow(i);
         cell = row.createCell(1);
         cell.setCellValue(resultSet.getInt("ANI"));
         cell = row.createCell(2);
         cell.setCellValue(resultSet.getString("Call Date Time"));
         cell = row.createCell(3);
         cell.setCellValue(resultSet.getString("Duration"));
         cell = row.createCell(4);
         cell.setCellValue(resultSet.getString("Id"));
         cell = row.createCell(5);
         cell.setCellValue(resultSet.getString("Short Code"));
         cell=row.createCell(6);
         cell.setCellValue(resultSet.getString("Portal"));
         cell=row.createCell(7);
         cell.setCellValue(resultSet.getString("Click Id"));
         i++;
      }

      FileOutputStream out = new FileOutputStream(new File(myFile));
      workbook.write(out);
      out.close();
      System.out.println("exceldatabase.xlsx written successfully");
   }
}
    

