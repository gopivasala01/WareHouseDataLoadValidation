package mainPackage;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Properties;

import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WareHouseDataLoadValidation {
    static final String CONNECTION_URL = "jdbc:sqlserver://azrsrv001.database.windows.net;databaseName=HomeRiverDB;user=service_sql02;password=xzqcoK7T;encrypt=true;trustServerCertificate=true;";
    static String FILE_PATH = "DataLoadsValidation.xlsx"; // Updated file path

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

        try {
            connection = DriverManager.getConnection(CONNECTION_URL);
            statement = connection.createStatement();
            executeSecondQuery(statement, resultSet);
            sendEmail();

        } finally {
            closeResources(resultSet, statement, connection);
        }
    }

    static void executeSecondQuery(Statement statement, ResultSet resultSet) throws SQLException, IOException {
        String secondQuery = "Select TableName, Company, JSONCount,StagingTableCount, ProdTableCount,JSONCountRetrievalDate\r\n"
                + "from WareHouseDataLoadValidation  \r\n"
                + "where Convert(date,JSONCountRetrievalDate) =Convert(date, getdate()-1) --order by TableName desc\r\n"
                + "and (JSONCount<>StagingTableCount or StagingTableCount<>ProdTableCount or ProdTableCount<>JSONCount)\r\n"
                + "order by TableName desc";

        resultSet = statement.executeQuery(secondQuery);

        XSSFWorkbook workbook = new XSSFWorkbook();
        org.apache.poi.ss.usermodel.Sheet sheet = workbook.createSheet("Data");

        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("TableName");
        headerRow.createCell(1).setCellValue("Company");
        headerRow.createCell(2).setCellValue("JSONCount");
        headerRow.createCell(3).setCellValue("StagingTableCount");
        headerRow.createCell(4).setCellValue("ProdTableCount");
        headerRow.createCell(5).setCellValue("JSONCountRetrievalDate");

        int rowNum = 1;
        while (resultSet.next()) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(resultSet.getString("TableName"));
            row.createCell(1).setCellValue(resultSet.getString("Company"));
            if (resultSet.getObject("JSONCount") != null) {
                row.createCell(2).setCellValue(resultSet.getInt("JSONCount"));
            } else {
                row.createCell(2).setCellValue("");
            }
            if (resultSet.getObject("StagingTableCount") != null) {
                row.createCell(3).setCellValue(resultSet.getInt("StagingTableCount"));
            } else {
                row.createCell(3).setCellValue("");
            }
            if (resultSet.getObject("ProdTableCount") != null) {
                row.createCell(4).setCellValue(resultSet.getInt("ProdTableCount"));
            } else {
                row.createCell(4).setCellValue("");
            }
            row.createCell(5).setCellValue(resultSet.getString("JSONCountRetrievalDate"));
        }

        // Constructing file name with current date
        SimpleDateFormat dateFormat = new SimpleDateFormat("MMddyyyy");
        String formattedDate = dateFormat.format(new java.util.Date());
        FILE_PATH = "DataLoadsValidation_" + formattedDate + ".xlsx";

        FileOutputStream fileOutputStream = new FileOutputStream(FILE_PATH);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();
    }

    static void sendEmail() throws MessagingException {
        String smtpHost = "smtp.office365.com";
        String smtpPort = "587";
        String emailFrom = "santosh.p@beetlerim.com";
        String emailTo = "gopi.v@beetlerim.com, ratna@beetlerim.com , santosh.p@beetlerim.com, dahoffman@homeriver.com";
        final String username = "santosh.p@beetlerim.com";
        final String password = "Welcome@123";

        Properties props = new Properties();
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.host", smtpHost);
        props.put("mail.smtp.port", smtpPort);
        props.put("mail.smtp.ssl.protocols", "TLSv1.2");

        Session session = Session.getInstance(props, new javax.mail.Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(username, password);
            }
        });

        // Format the date as "MM/dd/yyyy"
        SimpleDateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
        String formattedDate = dateFormat.format(new java.util.Date());

        Message message = new MimeMessage(session);
        message.setFrom(new InternetAddress(emailFrom));
        message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(emailTo));
        message.setSubject("Data Loads Validation for date: " + formattedDate);

        MimeBodyPart mimeBodyPart = new MimeBodyPart();
        try {
            mimeBodyPart.attachFile(FILE_PATH);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (MessagingException e) {
            e.printStackTrace();
        }

        MimeBodyPart messageBodyPart = new MimeBodyPart();
        try {
            messageBodyPart.setText("Hi All,\n\nPlease find the attachment for the count differences in the loads:\n\nRegards,\nHomeRiver Group.");
        } catch (MessagingException e) {
            e.printStackTrace();
        }

        Multipart multipart = new MimeMultipart();
        multipart.addBodyPart(mimeBodyPart);
        multipart.addBodyPart(messageBodyPart);

        message.setContent(multipart);

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
