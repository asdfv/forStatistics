package forAsteriskStats;

import java.sql.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

public class MissedFinder {

    public static String toStringMissNeedCallBack(List<Entry> list) {
        if (list.size() == 0) return "none";
        else {
            String result = "";
            for (Entry e : list) {
                if (e.getNeedCallBack()) {
                    result += e;
                }
            }
            if (result == "") return "none";
            else return result;
        }
    }

    private static final String FIND_MISSED = "select data2, time, 'missed' as status from queue_log where time like concat(?, '%') and event = 'ENTERQUEUE' and callid in (" +
            "select callid from queue_log where time like concat(?, '%') and event = 'ABANDON' and data3 > 15 );";

    private static final String FIND_ANSWERED = "select data2, time, 'answered' as status from queue_log where time like concat(?, '%') and event = 'ENTERQUEUE' and callid in (" +
            "select callid from queue_log where time like concat(?, '%') and event = 'CONNECT' );";

    private static final String FIND_OUTGOING = "SELECT dst, calldate, 'outgoing' as status FROM cdr WHERE calldate like concat(?, '%') and ( dstchannel LIKE 'DAHDI%' OR dstchannel LIKE 'SIP/dinstar-trunk-gsm%' );";

    public static void main(String[] args) {

        // Current date if needed
        long curTime = System.currentTimeMillis();
        String curDate = new SimpleDateFormat("yyyy-MM-dd").format(curTime);

        // Variable declaration
        String period = "2016-07-29";

        Connection con = null;
        PreparedStatement missedPrepS = null;
        PreparedStatement answeredPrepS = null;
        PreparedStatement outgoingPrepS = null;
        ResultSet missedResS = null;
        ResultSet answeredResS = null;
        ResultSet outgoingResS = null;

        int countOfMiss = 0;
        int countOfAns = 0;
        int countOfOut = 0;
        int percentOfMiss = 0;

        List<Entry> missedList = new ArrayList<Entry>();
        List<Entry> answeredList = new ArrayList<Entry>();
        List<Entry> outgoingList = new ArrayList<Entry>();

        try {
            // Getting connection
            con = new ConnectionMaker().getConnection();

            // Creating prepare statements
            missedPrepS = con.prepareStatement(FIND_MISSED);
            missedPrepS.setString(1, period);
            missedPrepS.setString(2, period);

            answeredPrepS = con.prepareStatement(FIND_ANSWERED);
            answeredPrepS.setString(1, period);
            answeredPrepS.setString(2, period);

            outgoingPrepS = con.prepareStatement(FIND_OUTGOING);
            outgoingPrepS.setString(1, period);

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

            countOfMiss = missedList.size();
            countOfAns = answeredList.size();
            countOfOut = outgoingList.size();
            percentOfMiss = (countOfMiss > 0) ? countOfMiss * 100 / (countOfAns + countOfMiss) : 0;

            if (con != null) con.close();

            System.out.println("Day " + period +
                    "\nPercent of missed: " + percentOfMiss + " %" +
                    "\nAnswered: " + countOfAns + "\nMissed: " + countOfMiss +
                    "\nOutgoing: " + countOfOut + "\nMissed need to call back: \n" + toStringMissNeedCallBack(missedList));

        } catch (SQLException e) {
            e.printStackTrace();
        } finally { //Closing writer stream, connection, PS and RS
            try {
                if (con != null) con.close();
                if (missedPrepS != null) missedPrepS.close();
                if (answeredPrepS != null) answeredPrepS.close();
                if (outgoingPrepS != null) outgoingPrepS.close();
                if (missedResS != null) missedResS.close();
                if (answeredResS != null) answeredResS.close();
                if (outgoingResS != null) outgoingResS.close();

            } catch (SQLException e) { e.printStackTrace(); }
        }
    }
}
