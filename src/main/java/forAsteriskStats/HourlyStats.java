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

public class HourlyStats {

    private final static String MISSED_BY_HOUR = "select count(event) as 'countMiss' from queue_log where time like concat(?,'%') and hour(time) = ? and event = 'ABANDON' and data3 > 15;";
    private final static String INCOMING_BY_HOUR = "select count(event) as 'countIn' from queue_log where time like concat(?,'%') and hour(time) = ? and event = 'ENTERQUEUE';";
    private final static String OUTGOING_BY_HOUR = "select count(*) as 'countOut' from cdr WHERE calldate like concat(?, '%') and ( dstchannel like 'DAHDI%' or dstchannel like 'SIP/dinstar-trunk-gsm%') and hour(calldate) = ?;";
    private final static String CALL_DURATION_BY_HOUR = "select avg(data2) as 'countDur' from queue_log where time like concat(?,'%') and hour(time) = ? and (event = 'COMPLETEAGENT' or event = 'COMPLETECALLER');";

    public static void main(String[] args) {

        String period = "2016-10-10";
        int startHourTime = 9;
        int endHourTime = 24;

        int hourTime = startHourTime;
        String pathToExportFileXLS = "D://statsByHours.xls";
        Connection connection = null;
        PreparedStatement missPrepS;
        PreparedStatement inPrepS;
        PreparedStatement outPrepS;
        PreparedStatement durPrepS;

        ResultSet missResS;
        ResultSet inResS;
        ResultSet outResS;
        ResultSet durResS;

        FileOutputStream fileOutputStream = null;
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();

        Row[] rowList = new Row[endHourTime - startHourTime + 2];

        try {
            // Init for Row-array
            for (int i = 0; i < endHourTime - startHourTime + 2; i++) {
                rowList[i] = sheet.createRow(i);
            }

            connection = new ConnectionMaker().getConnection();
            fileOutputStream = new FileOutputStream(pathToExportFileXLS);

            missPrepS = connection.prepareStatement(MISSED_BY_HOUR);
            inPrepS = connection.prepareStatement(INCOMING_BY_HOUR);
            outPrepS = connection.prepareStatement(OUTGOING_BY_HOUR);
            durPrepS = connection.prepareStatement(CALL_DURATION_BY_HOUR);

            // Row 1 (0 position in array)
            rowList[0].createCell(0).setCellValue("Stats for the day " + period);

            // Row 2
            rowList[1].createCell(0).setCellValue("Hour"); // Cell A2
            rowList[1].createCell(1).setCellValue("Count Missed"); // B2
            rowList[1].createCell(2).setCellValue("Count Incoming"); // C2
            rowList[1].createCell(3).setCellValue("Count Outgoing"); // D2
            rowList[1].createCell(4).setCellValue("Average duration of a call, sec"); // E2

            System.out.println("Missed");
            for (hourTime = startHourTime; hourTime < endHourTime; hourTime++) {
                missPrepS.setString(1, period);
                missPrepS.setInt(2, hourTime);
                missResS = missPrepS.executeQuery();

                missResS.first();

                rowList[hourTime - startHourTime + 2].createCell(0).setCellValue(hourTime); // Cell A3, A4, A5...
                rowList[hourTime - startHourTime + 2].createCell(1).setCellValue(missResS.getInt("countMiss")); // B3, B4, B5...

                System.out.println(hourTime + " - " + missResS.getInt("countMiss"));
            }

            System.out.println("Incoming");
            for (hourTime = startHourTime; hourTime < endHourTime; hourTime++) {
                inPrepS.setString(1, period);
                inPrepS.setInt(2, hourTime);
                inResS = inPrepS.executeQuery();
                inResS.first();

                rowList[hourTime - startHourTime + 2].createCell(2).setCellValue(inResS.getInt("countIn")); // C3, C4, C5...

                System.out.println(hourTime + " - " + inResS.getInt("countIn"));
            }

            System.out.println("Outgoing");
            for (hourTime = startHourTime; hourTime < endHourTime; hourTime++) {
                outPrepS.setString(1, period);
                outPrepS.setInt(2, hourTime);
                outResS = outPrepS.executeQuery();
                outResS.first();

                rowList[hourTime - startHourTime + 2].createCell(3).setCellValue(outResS.getInt("countOut")); // D3, D4, D5...

                System.out.println(hourTime + " - " + outResS.getInt("countOut"));
            }

            System.out.println("Calls duration");
            for (hourTime = startHourTime; hourTime < endHourTime; hourTime++) {
                durPrepS.setString(1, period);
                durPrepS.setInt(2, hourTime);
                durResS = durPrepS.executeQuery();
                durResS.first();

                rowList[hourTime - startHourTime + 2].createCell(4).setCellValue(durResS.getInt("countDur")); // E - cells

                System.out.println(hourTime + " - " + durResS.getInt("countDur"));
            }

            workbook.write(fileOutputStream);

            System.out.println("Data has been exported to " + pathToExportFileXLS);

        } catch (SQLException e) { e.printStackTrace();
        } catch (IOException e) { e.printStackTrace();
        } finally { try {
            if (connection != null)  connection.close();
            if (fileOutputStream != null) fileOutputStream.close();
            if (workbook != null) workbook.close();
        }
            catch (SQLException e) { e.printStackTrace(); }
            catch (IOException e) { e.printStackTrace(); }
        }
    }
}
