package forAsteriskStats;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

public class ByMonthStats {
    private final static String MISSED_COUNT = "select count(event) as 'countMiss' from queue_log where time like concat(?,'%') and event = 'ABANDON' and data3 > 15;";
    private final static String INCOMING_COUNT = "select count(event) as 'countIn' from queue_log where time like concat(?,'%') and event = 'ENTERQUEUE';";
    private final static String ANSWERED_COUNT = "select count(event) as 'countAns' from queue_log where time like concat(?, '%') and event = 'CONNECT';";
    private final static String AVG_WAIT_TIME = "select avg(data1) as 'avgWaitTime' from queue_log where time like concat(?,'%') and event = 'CONNECT';";

    public static void main(String[] args) {

        String period = "2016";
        String[] months = {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"};
        String pathToExportFile = "D://ByMonth.xls";
        Connection connection = null;
        PreparedStatement missPrepS;
        PreparedStatement inPrepS;
        PreparedStatement ansPrepS;
        PreparedStatement waitPrepS;

        ResultSet missResS;
        ResultSet inResS;
        ResultSet ansResS;
        ResultSet waitResS;

        Workbook workbook;
        FileOutputStream fileOutputStream = null;

        workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();

        Row[] rowList = new Row[months.length + 2];

        // Init for row
        for (int i = 0; i < months.length + 2; i++) {
            rowList[i] = sheet.createRow(i);
        }

        try {
            connection = new ConnectionMaker().getConnection();
            fileOutputStream = new FileOutputStream(pathToExportFile);

            missPrepS = connection.prepareStatement(MISSED_COUNT);
            inPrepS = connection.prepareStatement(INCOMING_COUNT);
            ansPrepS = connection.prepareStatement(ANSWERED_COUNT);
            waitPrepS = connection.prepareStatement(AVG_WAIT_TIME);


            // Cell A1
            rowList[0].createCell(0).setCellValue(period);

            // Cells A2 - B2
            rowList[1].createCell(0).setCellValue("Month");
            rowList[1].createCell(1).setCellValue("Count incoming");
            rowList[1].createCell(2).setCellValue("Answered");
            rowList[1].createCell(3).setCellValue("Missed");
            rowList[1].createCell(4).setCellValue("Percent of miss");
            rowList[1].createCell(5).setCellValue("Avg wait time");

            for (String month : months) {
                int percentOfMiss;
                int countOfMiss;
                int countOfIn;
                int countOfAns;
                int avgWaitTime;

                missPrepS.setString(1, period + "-" + month);
                ansPrepS.setString(1, period + "-" + month);
                inPrepS.setString(1, period + "-" + month);
                waitPrepS.setString(1, period + "-" + month);

                missResS = missPrepS.executeQuery();
                inResS = inPrepS.executeQuery();
                ansResS = ansPrepS.executeQuery();
                waitResS = waitPrepS.executeQuery();

                missResS.first();
                inResS.first();
                ansResS.first();
                waitResS.first();

                countOfMiss = missResS.getInt("countMiss");
                countOfIn = inResS.getInt("countIn");
                countOfAns = ansResS.getInt("countAns");
                avgWaitTime = waitResS.getInt("avgWaitTime");

                percentOfMiss = countOfIn > 0 ? countOfMiss * 100 / (countOfMiss + countOfAns) : 0;

                // Cells A3 - B33
                rowList[Integer.parseInt(month) + 1].createCell(0).setCellValue(Integer.parseInt(month));
                rowList[Integer.parseInt(month) + 1].createCell(1).setCellValue(countOfIn);
                rowList[Integer.parseInt(month) + 1].createCell(2).setCellValue(countOfAns);
                rowList[Integer.parseInt(month) + 1].createCell(3).setCellValue(countOfMiss);
                rowList[Integer.parseInt(month) + 1].createCell(4).setCellValue(percentOfMiss);
                rowList[Integer.parseInt(month) + 1].createCell(5).setCellValue(avgWaitTime);

                System.out.println(month + " - " + percentOfMiss + " - " + countOfAns + " - " + avgWaitTime);
            }
            workbook.write(fileOutputStream);
            System.out.println("Data for " + period + " has been exported to " + pathToExportFile);

        } catch (SQLException e) { e.printStackTrace();
        } catch (IOException e) { e.printStackTrace();
        } finally {
            try {
                if (connection != null) connection.close();
                if (workbook != null) workbook.close();
                if (fileOutputStream != null) fileOutputStream.close();
            }
            catch (SQLException e) { e.printStackTrace(); }
            catch (IOException e) { e.printStackTrace(); }
        }
    }
}
