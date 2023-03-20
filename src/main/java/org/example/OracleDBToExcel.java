package org.example;

import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class OracleDBToExcel {
    public static void main(String[] args) throws Exception {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Result");

        Connection conn = DriverManager.getConnection("jdbc:oracle:thin:@//localhost:1521/xe", "username", "password");
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery("SELECT * FROM table");

        int rownum = 0;
        while(rs.next()) {
            Row row = sheet.createRow(rownum++);
            for(int i = 1; i <= rs.getMetaData().getColumnCount(); i++) {
                Cell cell = row.createCell(i-1);
                cell.setCellValue(rs.getString(i));
            }
        }
        rs.close();
        stmt.close();
        conn.close();

        FileOutputStream outputStream = new FileOutputStream("result.xlsx");
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }
}
