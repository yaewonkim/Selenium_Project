package PDFEmail;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class BaseClass{

    protected static WebDriver driver;

	/* 
	 * WebDriver 생성
	 */
    public static WebDriver getDriver(){

        if(driver==null){
    	System.setProperty("webdriver.gecko.driver", "C:\\Users\\YaewonKim\\workspace\\seleniumTest\\drivers\\geckodriver.exe");
    	driver = new FirefoxDriver();
        }

        return driver;
    }


    /**
     * 스크린 샷 찍기
     * @param webdriver
     * @param fileWithPath
     * @throws Exception
     */
    public static void takeSnapShot(WebDriver webdriver,String fileWithPath) throws Exception{
    	 	System.out.println("screenshot file path: "+fileWithPath);
    		File SrcFile=((TakesScreenshot)webdriver).getScreenshotAs(OutputType.FILE);

            //Move image file to new destination

            File DestFile=new File(fileWithPath);

            //Copy file at destination

            FileUtils.copyFile(SrcFile, DestFile);

    }
    

	/**
	 * Gmail 전송하기
	 * @param from
	 * @param pass
	 * @param to
	 * @param subject
	 * @param body
	 * @param suiteName
	 */
	public static void sendPDFReportByGMail(String from, String pass, String to, String subject, String body, String suiteName) {
	
       Properties props = System.getProperties();
       String host = "smtp.gmail.com";
       props.put("mail.smtp.starttls.enable", "true");
       props.put("mail.smtp.host", host);
       props.put("mail.smtp.user", from);
       props.put("mail.smtp.password", pass);
       props.put("mail.smtp.port", "587");
       props.put("mail.smtp.auth", "true");

       Session session = Session.getDefaultInstance(props);
       MimeMessage message = new MimeMessage(session);

       try {
          //Set from address
           message.setFrom(new InternetAddress(from));
            message.addRecipient(Message.RecipientType.TO, new InternetAddress(to));
          //Set subject
           message.setSubject(subject);
           message.setText(body);
         
           BodyPart objMessageBodyPart = new MimeBodyPart();
           
           objMessageBodyPart.setText(suiteName+"테스트를 진행한 결과입니다.");
           
           Multipart multipart = new MimeMultipart();

           multipart.addBodyPart(objMessageBodyPart);

           objMessageBodyPart = new MimeBodyPart();

           //Set path to the pdf report file
           String filename = System.getProperty("user.dir")+"\\report\\"+suiteName+".pdf"; 
           System.out.println("pdf filepath: "+filename);
           
           //Create data source to attach the file in mail
           DataSource source = new FileDataSource(filename);
           
           
           objMessageBodyPart.setDataHandler(new DataHandler(source));
           objMessageBodyPart.setFileName(suiteName+".pdf");

           multipart.addBodyPart(objMessageBodyPart);

           message.setContent(multipart);
           Transport transport = session.getTransport("smtp");
           transport.connect(host, from, pass);
           transport.sendMessage(message, message.getAllRecipients());
           transport.close();
       }
       catch (AddressException ae) {
           ae.printStackTrace();
       }
       catch (MessagingException me) {
           me.printStackTrace();
       }
   }
	
	
	
	/**
	 * 엑셀시트로부터 데이터 가져오기
	 * @param File Name
	 * @param Sheet Name
	 * @return
	 */
	public String[][] getExcelData(String fileName, String sheetName) {
		String[][] arrayExcelData = null;
		try {
			FileInputStream fs = new FileInputStream(fileName);
			Workbook wb = Workbook.getWorkbook(fs);
			Sheet sh = wb.getSheet(sheetName);

			int totalNoOfCols = sh.getColumns();
			int totalNoOfRows = sh.getRows();
			
			arrayExcelData = new String[totalNoOfRows-1][totalNoOfCols];
			
			for (int i= 1 ; i < totalNoOfRows; i++) {

				for (int j=0; j < totalNoOfCols; j++) {
					arrayExcelData[i-1][j] = sh.getCell(j, i).getContents();
				}

			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
			e.printStackTrace();
		} catch (BiffException e) {
			e.printStackTrace();
		}
		return arrayExcelData;
	}

	

}