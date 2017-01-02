package forAsteriskStats;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.Properties;

public class ConnectionMaker {

    private static Connection connection = null;
    Properties properties = new Properties();

    public ConnectionMaker() {
        try {
            properties.load(new FileInputStream("/db.properties"));
            connection = DriverManager.getConnection(properties.getProperty("URL"), properties.getProperty("USERNAME"), properties.getProperty("PASSWORD"));
        }
        catch (IOException e) { e.printStackTrace(); }
        catch (SQLException e) { e.printStackTrace(); }
    }

    public static Connection getConnection() {
        return connection;
    }
}
