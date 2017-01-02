package forAsteriskStats;

import java.sql.Timestamp;
import java.text.SimpleDateFormat;

public class Entry {
    long number;
    Timestamp calldate;
    boolean needCallBack = true; // Call need a call back, when this not get answered or calling back later

    public Entry(long number, Timestamp calldate) {
        this.number = number;
        this.calldate = calldate;
    }

    public long getNumber() {
        return number;
    }

    public Timestamp getCalldate() {
        return calldate;
    }

    public void setNumber(long number) {
        this.number = number;
    }

    public void setCalldate(Timestamp calldate) {
        this.calldate = calldate;
    }

    public boolean getNeedCallBack() {
        return needCallBack;
    }

    public void setNeedCallBack(boolean needCallBack) {
        this.needCallBack = needCallBack;
    }

    @Override
    public String toString() {
        return number + " - " + new SimpleDateFormat("HH:mm:ss").format(calldate) + "\n";
    }
}
