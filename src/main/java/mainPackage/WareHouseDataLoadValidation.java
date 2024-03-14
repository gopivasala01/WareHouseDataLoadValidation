package mainPackage;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;


import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WareHouseDataLoadValidation {
    // JDBC URL, username, and password of SQL server
    static final String CONNECTION_URL = "jdbc:sqlserver://azrsrv001.database.windows.net;databaseName=HomeRiverDB;user=service_sql02;password=xzqcoK7T;encrypt=true;trustServerCertificate=true;";
    static final String FILE_PATH = "output.xlsx";

    public static void main(String[] args) {
        try {
            executeQueriesAndSendEmail();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static void executeQueriesAndSendEmail() throws Exception {
        Connection connection = null;
        Statement statement = null;
        ResultSet resultSet = null;

        boolean firstQueryHasData = false; // Flag to indicate if the first query has data

        try {
            // Establishing a connection to the database
            connection = DriverManager.getConnection(CONNECTION_URL);

            // Creating a statement
            statement = connection.createStatement();

            // Execute queries
            firstQueryHasData = executeFirstQuery(statement, resultSet);

            if (firstQueryHasData) {
                executeSecondQuery(statement, resultSet);
                sendEmail();
            }

        } finally {
            // Closing the resources
            closeResources(resultSet, statement, connection);
        }

        // If the first query did not return any data, return from the main method
        if (!firstQueryHasData) {
            System.out.println("No data found for the first query. Exiting.");
        }
    }

    static boolean executeFirstQuery(Statement statement, ResultSet resultSet) throws SQLException {
        // First query
        String firstQuery = "SELECT * FROM WareHouseDataLoadValidation WHERE CONVERT(date, JSONCountRetrievalDate) = CONVERT(date, GETDATE()) AND JSONCount IS NOT NULL AND StagingTableCount IS NOT NULL AND ProdTableCount IS NOT NULL";
        resultSet = statement.executeQuery(firstQuery);

        // Checking if the result set has more than 0 rows
        return resultSet.isBeforeFirst();
    }

    static void executeSecondQuery(Statement statement, ResultSet resultSet) throws SQLException, IOException {
        // Second query
        String secondQuery = "SELECT TableName, Company, JSONCount, StagingTableCount, ProdTableCount, JSONCountRetrievalDate " +
                "FROM WareHouseDataLoadValidation " +
                "WHERE CONVERT(date, JSONCountRetrievalDate) = CONVERT(date, GETDATE()) " + 
                "AND JSONCount <> StagingTableCount AND StagingTableCount <> ProdTableCount AND ProdTableCount <> JSONCount " +
                "ORDER BY TableName";

        resultSet = statement.executeQuery(secondQuery);

        // Create a workbook and a sheet
        XSSFWorkbook workbook = new XSSFWorkbook();
        org.apache.poi.ss.usermodel.Sheet sheet = workbook.createSheet("Data");

        // Create header row
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("TableName");
        headerRow.createCell(1).setCellValue("Company");
        headerRow.createCell(2).setCellValue("JSONCount");
        headerRow.createCell(3).setCellValue("StagingTableCount");
        headerRow.createCell(4).setCellValue("ProdTableCount");
        headerRow.createCell(5).setCellValue("JSONCountRetrievalDate");

        int rowNum = 1;
        // Iterate over the result set and populate the sheet
     // Iterate over the result set and populate the sheet
        while (resultSet.next()) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(resultSet.getString("TableName"));
            row.createCell(1).setCellValue(resultSet.getString("Company"));
            if (resultSet.getObject("JSONCount") != null) {
                row.createCell(2).setCellValue(resultSet.getInt("JSONCount"));
            } else {
                row.createCell(2).setCellValue(""); // Set empty string instead of 0
            }
            if (resultSet.getObject("StagingTableCount") != null) {
                row.createCell(3).setCellValue(resultSet.getInt("StagingTableCount"));
            } else {
                row.createCell(3).setCellValue(""); // Set empty string instead of 0
            }
            if (resultSet.getObject("ProdTableCount") != null) {
                row.createCell(4).setCellValue(resultSet.getInt("ProdTableCount"));
            } else {
                row.createCell(4).setCellValue(""); // Set empty string instead of 0
            }
            row.createCell(5).setCellValue(resultSet.getString("JSONCountRetrievalDate"));
        }

        // Write the workbook to a file
        FileOutputStream fileOutputStream = new FileOutputStream(FILE_PATH);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();
    }

    static void sendEmail() throws MessagingException {
        // SMTP configuration
        String smtpHost = "smtp.office365.com";
        String smtpPort = "587";
        String emailFrom = "santosh.p@beetlerim.com";
        String emailTo = "gopi.v@beetlerim.com, ratna@beetlerim.com , santosh.p@beetlerim.com";

        // Sender's credentials
        final String username = "santosh.p@beetlerim.com";
        final String password = "Welcome@123";

        // Email properties
        Properties props = new Properties();
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.host", smtpHost);
        props.put("mail.smtp.port", smtpPort);
        props.put("mail.smtp.ssl.protocols", "TLSv1.2");

        // Create a Session object
        Session session = Session.getInstance(props, new javax.mail.Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(username, password);
            }
        });

        // Create a MimeMessage object
        Message message = new MimeMessage(session);
        message.setFrom(new InternetAddress(emailFrom));
        message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(emailTo));
        message.setSubject("Data from SQL Server");

        // Create MimeBodyPart and attach the Excel file
        MimeBodyPart mimeBodyPart = new MimeBodyPart();
        try {
            mimeBodyPart.attachFile(FILE_PATH);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (MessagingException e) {
            e.printStackTrace();
        }

        // Create Multipart object and add MimeBodyPart objects to it
        Multipart multipart = new MimeMultipart();
        multipart.addBodyPart(mimeBodyPart);

        // Set the content of the email
        message.setContent(multipart);

        // Send the email
        session.setDebug(true);
        Transport.send(message);

        System.out.println("Email sent successfully!");
    }

    static void closeResources(ResultSet resultSet, Statement statement, Connection connection) {
        try {
            if (resultSet != null) resultSet.close();
            if (statement != null) statement.close();
            if (connection != null) connection.close();
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }
}