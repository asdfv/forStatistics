package forAsteriskStats;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Test {

    public static void main(String[] args) {
        Workbook workbook = null;
        FileOutputStream fileOutputStream = null;

        try {

            fileOutputStream = new FileOutputStream("D://workbook.xls");
            workbook = new HSSFWorkbook();
            Sheet sheet = workbook.createSheet();
            Row row = sheet.createRow(0);
            Cell cell;

            row.createCell(0).setCellValue("First column");
            row.createCell(1).setCellValue("Second column");

            List<Row> rowList = new ArrayList<Row>();

            for (int i = 1; i < 10; i++) {
                Row row1 = sheet.createRow(i);
                rowList.add(row1);
                cell = row1.createCell(0);
                cell.setCellValue(i);
            }

            for (int i = 1; i < 10; i++) {


                cell = rowList.get(i - 1).createCell(1);
                cell.setCellValue(i + " - 2");
            }

            workbook.write(fileOutputStream);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
                fileOutputStream.close();
            } catch (IOException e) { e.printStackTrace(); }
        }
    }

}
