package forAsteriskStats;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;

public class CountMissNotCallBackByDay {

    private static final String FIND_MISSED = "select data2, time, 'missed' as status from queue_log where time like concat(?, '%') and event = 'ENTERQUEUE' and callid in (" +
            "select callid from queue_log where time like concat(?, '%') and event = 'ABANDON' and data3 > 15 );";

    private static final String FIND_ANSWERED = "select data2, time, 'answered' as status from queue_log where time like concat(?, '%') and event = 'ENTERQUEUE' and callid in (" +
            "select callid from queue_log where time like concat(?, '%') and event = 'CONNECT' );";

    private static final String FIND_OUTGOING = "SELECT dst, calldate, 'outgoing' as status FROM cdr WHERE calldate like concat(?, '%') and ( dstchannel LIKE 'DAHDI%' OR dstchannel LIKE 'SIP/dinstar-trunk-gsm%' );";

    public static void main(String[] args) {

        long startTimeMs = System.currentTimeMillis();

        // Variable declaration
        int[] noCallBackMissList = new int[31];
        String[] days = {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31"};
        String period = "2016-12";
        String periodDay;
        String path = "D://MissedNotCallBackByDay " + period + ".xls";

        Connection con = null;
        PreparedStatement missedPrepS = null;
        PreparedStatement answeredPrepS = null;
        PreparedStatement outgoingPrepS = null;
        ResultSet missedResS = null;
        ResultSet answeredResS = null;
        ResultSet outgoingResS = null;

        FileOutputStream fileOutputStream = null;

        List<Entry> missedList = new ArrayList<Entry>();
        List<Entry> answeredList = new ArrayList<Entry>();
        List<Entry> outgoingList = new ArrayList<Entry>();

        // Create workbook and init for rows
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        Row[] rowList = new Row[days.length + 2];

        for (int i = 0; i < days.length + 2; i++) {
            rowList[i] = sheet.createRow(i);
        }

        try {
            // Getting connection
            con = new ConnectionMaker().getConnection();
            fileOutputStream = new FileOutputStream(path);

            for (int i = 0; i < days.length; i++) {

                missedList.clear();
                answeredList.clear();
                outgoingList.clear();

                // Creating prepare statements
                periodDay = period + "-" + days[i];

                missedPrepS = con.prepareStatement(FIND_MISSED);
                missedPrepS.setString(1, periodDay);
                missedPrepS.setString(2, periodDay);

                answeredPrepS = con.prepareStatement(FIND_ANSWERED);
                answeredPrepS.setString(1, periodDay);
                answeredPrepS.setString(2, periodDay);

                outgoingPrepS = con.prepareStatement(FIND_OUTGOING);
                outgoingPrepS.setString(1, periodDay);

                // Getting resutSets
                missedResS = missedPrepS.executeQuery();
                answeredResS = answeredPrepS.executeQuery();
                outgoingResS = outgoingPrepS.executeQuery();

                // ResultSets to arrays of instance Entry class
                while (missedResS.next()) {
                    try {
                        missedResS.getLong("data2"); // data2 field can be "Anonymous"
                    } catch (SQLException e) { continue; }

                    Entry miss = new Entry(missedResS.getLong("data2"), missedResS.getTimestamp("time"));
                    missedList.add(miss);
                }

                while (answeredResS.next()) {
                    try {
                        answeredResS.getLong("data2"); // data2 field can be "Anonymous"
                    } catch (SQLException e) { continue; }

                    Entry ans = new Entry(answeredResS.getLong("data2"), answeredResS.getTimestamp("time"));
                    answeredList.add(ans);
                }
                while (outgoingResS.next()) {
                    try {
                        outgoingResS.getLong("dst"); // dst can be bigger than long-type in the case of an incorrect dialing
                    } catch (SQLException e) { continue; }
                    Entry out = new Entry(outgoingResS.getLong("dst"), outgoingResS.getTimestamp("calldate"));
                    outgoingList.add(out);
                }

                // Comparison of fetches by last digits in numbers

                for (Entry miss : missedList) {
                    long fullNumber = miss.getNumber();
                    Timestamp dateMiss = miss.getCalldate();
                    long countOfDigit = 1000000; // The number of zeros - is the number of last digits
                    long number = fullNumber % countOfDigit;

                    // Set status NeedCallBack - false for later answered calls
                    for (Entry ans : answeredList) {
                        if ((ans.getNumber() % countOfDigit == number) && (ans.getCalldate().after(dateMiss))) {
                            miss.setNeedCallBack(false);
                        }
                    }

                    // Set status NeedCallBack - false for later outgoing calls
                    for (Entry out : outgoingList) {
                        if ((out.getNumber() % countOfDigit == number) && (out.getCalldate().after(dateMiss))) {
                            miss.setNeedCallBack(false);
                        }
                    }
                }

                // Counting the number of missed with status "need call back"
                int countOfNotAnsweredMiss = 0;

                for (Entry e : missedList) {
                    if (e.getNeedCallBack()) countOfNotAnsweredMiss++;
                }

                noCallBackMissList[i] = countOfNotAnsweredMiss;
            }

            // Out noCallBackMissList to console and output xls
            rowList[0].createCell(0).setCellValue("Day");
            rowList[0].createCell(1).setCellValue("Count of missed without call back");

            for (int k = 0; k < days.length; k++) {
                System.out.println(period + "-" + days[k] + " - " + noCallBackMissList[k]);
                rowList[k + 1].createCell(0).setCellValue(days[k]);
                rowList[k + 1].createCell(1).setCellValue(noCallBackMissList[k]);
            }

            workbook.write(fileOutputStream);

        } catch (SQLException e) { e.printStackTrace();
        } catch (FileNotFoundException e) { e.printStackTrace();
        } catch (IOException e) { e.printStackTrace();
        } finally { //Closing writer stream, connection, PS and RS
            try {
                if (con != null) con.close();
                if (missedPrepS != null) missedPrepS.close();
                if (answeredPrepS != null) answeredPrepS.close();
                if (outgoingPrepS != null) outgoingPrepS.close();
                if (missedResS != null) missedResS.close();
                if (answeredResS != null) answeredResS.close();
                if (outgoingResS != null) outgoingResS.close();
                if (workbook != null) workbook.close();
                if (fileOutputStream != null) fileOutputStream.close();
            } catch (SQLException e) { e.printStackTrace(); }
            catch (IOException e) { e.printStackTrace(); }
        }
        long endTimeMs = System.currentTimeMillis();
        System.out.println("Exe-time: " + (endTimeMs - startTimeMs));
    }

}
