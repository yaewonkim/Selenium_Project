package PDFEmail;

import java.awt.Color;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Set;

import org.testng.ITestContext;
import org.testng.ITestListener;
import org.testng.ITestResult;

import com.lowagie.text.Chunk;
import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Element;
import com.lowagie.text.Font;
import com.lowagie.text.FontFactory;
import com.lowagie.text.Image;
import com.lowagie.text.Paragraph;
import com.lowagie.text.pdf.PdfAction;
import com.lowagie.text.pdf.PdfPCell;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;

public class JypersionListener implements ITestListener {
	/**
	 * Document
	 */
	private Document document = null;
	
	/**
	 * PdfPTables
	 */
	PdfPTable successTable = null, failTable = null;
	
	/**
	 * throwableMap
	 */
	private HashMap<Integer, Throwable> throwableMap = null;
	
	/**
	 * nbExceptions
	 */
	private int nbExceptions = 0;
	
	//한국어 처리 시도중
	public static final String KOREAN = "\ube48\uc9d1";
	
	private int FailedTestNum = 0;
	private int SuccessTestNum = 0;
	
	private Paragraph P_SuccessTestNum = null;
	private Paragraph P_FailedTestNum = null;
	
	/**
	 * JyperionListener
	 */
	public JypersionListener() {
		log("JyperionListener() start");
		
		this.document = new Document();
		
		this.throwableMap = new HashMap<Integer, Throwable>();
	}
	
	
	/* 
	 * test case 시작
	 */
	public void onStart(ITestContext context) {
		log(("Start Of Test-Suite Execution->"+context.getName()));
		
		try {
			String filepath = System.getProperty("user.dir")+"\\report\\"; 
			PdfWriter.getInstance(this.document, new FileOutputStream(filepath+context.getName()+".pdf"));
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		this.document.open();
		
//		한국어 처리 시도중
//		Paragraph title = new Paragraph(String.format(context.getName()+"테스트결과 TEST CASE", KOREAN),
//				FontFactory.getFont(FontFactory.HELVETICA, 20, Font.BOLD, new Color(0, 0, 0)));
		
		Paragraph title = new Paragraph("TEST Result\n\n",
				FontFactory.getFont(FontFactory.HELVETICA, 20, Font.BOLD, new Color(0, 0, 0)));
		title.setAlignment(Element.ALIGN_CENTER);
		
		
		try {
			this.document.add(title);
			
			//날짜 형식 설정
			SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
			Paragraph date = new Paragraph("Test Name: "+context.getName()+"\n"+
										   "Date: "+dateFormat.format(new Date()).toString()+"\n\n");
			
			this.document.add(date);
			
		} catch (DocumentException e1) {
			e1.printStackTrace();
		}
	}
	
	
	
	/* 
	 * test case 성공
	 */
	public void onTestSuccess(ITestResult result) {
		SuccessTestNum++;
		log("<<Success>> "+result.getName());
		
		//'test명_날짜_시간'으로 파일명 생성
	    Date today = new Date();
				        
		SimpleDateFormat date = new SimpleDateFormat("MMdd");
		SimpleDateFormat time = new SimpleDateFormat("hhmmss");
		
	    String screenshotName=result.getName()+"_"+date.format(today)+"_"+time.format(today)+".png";
	
	    String screenshot=System.getProperty("user.dir")+"\\"+"screenshot\\"+screenshotName;
	    try {
			BaseClass.takeSnapShot(BaseClass.getDriver(), screenshot);
		} catch (Exception e) {
			e.printStackTrace();
		}
	    
		if (successTable == null) {
			this.successTable = new PdfPTable(new float[]{.3f, .2f, .1f, .4f});
			
			Paragraph p = new Paragraph("PASSED TESTS", new Font(Font.TIMES_ROMAN, Font.DEFAULTSIZE, Font.BOLD));
			p.setAlignment(Element.ALIGN_CENTER);
			PdfPCell cell = new PdfPCell(p);
			
			cell.setColspan(4);
			cell.setBackgroundColor(Color.GREEN);
			this.successTable.addCell(cell);
			
			P_SuccessTestNum = new Paragraph();
			cell = new PdfPCell(P_SuccessTestNum);
			cell.setColspan(4);
			this.successTable.addCell(cell);
			
			cell = new PdfPCell(new Paragraph("Class"));
			cell.setBackgroundColor(Color.LIGHT_GRAY);
			this.successTable.addCell(cell);
			cell = new PdfPCell(new Paragraph("Method"));
			cell.setBackgroundColor(Color.LIGHT_GRAY);
			this.successTable.addCell(cell);
			cell = new PdfPCell(new Paragraph("Time (ms)"));
			cell.setBackgroundColor(Color.LIGHT_GRAY);
			this.successTable.addCell(cell);
			cell = new PdfPCell(new Paragraph("Exception"));
			cell.setBackgroundColor(Color.LIGHT_GRAY);
			this.successTable.addCell(cell);
		}
		
		PdfPCell cell = new PdfPCell(new Paragraph(result.getTestClass().toString()));
		this.successTable.addCell(cell);
		cell = new PdfPCell(new Paragraph(result.getMethod().getMethodName().toString()));
		this.successTable.addCell(cell);
		cell = new PdfPCell(new Paragraph("" + (result.getEndMillis()-result.getStartMillis())));
		this.successTable.addCell(cell);

		Throwable throwable = result.getThrowable();
		if (throwable != null) {
			this.throwableMap.put(new Integer(throwable.hashCode()), throwable);
			this.nbExceptions++;
			
			String str_throwable = throwable.getClass().toString();
			str_throwable = str_throwable.replace("class ", "");
			
			Paragraph excep = new Paragraph(
					new Chunk(str_throwable, 
							new Font(Font.TIMES_ROMAN, Font.DEFAULTSIZE, Font.UNDERLINE)).
							setLocalGoto("" + throwable.hashCode()));
			
	     	//스크린샷 클릭 시 확대됨
		    Chunk imdb = new Chunk("[SCREEN SHOT]", new Font(Font.TIMES_ROMAN, Font.DEFAULTSIZE, Font.UNDERLINE));
	        imdb.setAction(new PdfAction("file:///"+screenshot)); 
	        
	        excep.add(imdb);
	        
			cell = new PdfPCell(excep);
			this.successTable.addCell(cell);
	
		} else {
			
			Paragraph p = new Paragraph();

			//스크린샷 클릭 시 확대됨
		    Chunk imdb = new Chunk("[SCREEN SHOT]", new Font(Font.TIMES_ROMAN, Font.DEFAULTSIZE, Font.UNDERLINE));
	        imdb.setAction(new PdfAction("file:///"+screenshot)); 
	        
	        p.add(imdb);
			
			this.successTable.addCell(new PdfPCell(p));
		}  //end if
		
		// PDF에 에러 screenshot 넣기
		Image image = null;
		try {
			image = Image.getInstance(screenshot);
			
			image.scaleToFit(500,300);
			this.document.add(image);
			
		} catch (Exception e) {
			e.printStackTrace();
		} 
		
	}

	
	/* 
	 * test case 실패
	 */
	public void onTestFailure(ITestResult result) {
		log("<<Failed>> "+result.getName());
		FailedTestNum++;
		
		//'test명_날짜_시간'으로 파일명 생성
		Date today = new Date();
		        
		SimpleDateFormat date = new SimpleDateFormat("MMdd");
	    SimpleDateFormat time = new SimpleDateFormat("hhmmss");
		        
	    String screenshotName=result.getName()+"_"+date.format(today)+"_"+time.format(today)+".png";
			
		String screenshot=System.getProperty("user.dir")+"\\"+"screenshot\\"+screenshotName;
		try {
			BaseClass.takeSnapShot(BaseClass.getDriver(), screenshot);
		} catch (Exception e) {
			e.printStackTrace();
		}
		if (this.failTable == null) {
			this.failTable = new PdfPTable(new float[]{.3f, .2f, .1f, .4f});
			this.failTable.setTotalWidth(20f);
			Paragraph p = new Paragraph("FAILED TESTS", new Font(Font.TIMES_ROMAN, Font.DEFAULTSIZE, Font.BOLD));
			p.setAlignment(Element.ALIGN_CENTER);
			PdfPCell cell = new PdfPCell(p);
			cell.setColspan(4);
			cell.setBackgroundColor(Color.RED);
			this.failTable.addCell(cell);
			
			P_FailedTestNum = new Paragraph();
			cell = new PdfPCell(P_FailedTestNum);
			cell.setColspan(4);
			this.failTable.addCell(cell);
			
			cell = new PdfPCell(new Paragraph("Class"));
			cell.setBackgroundColor(Color.LIGHT_GRAY);
			this.failTable.addCell(cell);
			cell = new PdfPCell(new Paragraph("Method"));
			cell.setBackgroundColor(Color.LIGHT_GRAY);
			this.failTable.addCell(cell);
			cell = new PdfPCell(new Paragraph("Time (ms)"));
			cell.setBackgroundColor(Color.LIGHT_GRAY);
			this.failTable.addCell(cell);
			cell = new PdfPCell(new Paragraph("Exception"));
			cell.setBackgroundColor(Color.LIGHT_GRAY);
			this.failTable.addCell(cell);
		}
		
		PdfPCell cell = new PdfPCell(new Paragraph(result.getTestClass().toString()));
		this.failTable.addCell(cell);
		cell = new PdfPCell(new Paragraph(result.getMethod().getMethodName().toString()));
		this.failTable.addCell(cell);
		cell = new PdfPCell(new Paragraph("" + (result.getEndMillis()-result.getStartMillis())));
		this.failTable.addCell(cell);
		
		Throwable throwable = result.getThrowable();
		if (throwable != null) {
			
			this.throwableMap.put(new Integer(throwable.hashCode()), throwable);
		
			this.nbExceptions++;
			
			String str_throwable = throwable.getClass().toString();
			str_throwable = str_throwable.replace("class ", "");
			
		    Paragraph  excep = new Paragraph(str_throwable+"\n");
		    
		    //스크린샷 클릭 시 확대됨
		    Chunk imdb = new Chunk("[SCREEN SHOT]", new Font(Font.TIMES_ROMAN, Font.DEFAULTSIZE, Font.UNDERLINE));
	        imdb.setAction(new PdfAction("file:///"+screenshot)); 
	        
	        excep.add(imdb);
	
			cell = new PdfPCell(excep);
			this.failTable.addCell(cell);
			
		} else {
			this.failTable.addCell(new PdfPCell(new Paragraph("")));
		}  //end if
		
		// PDF에 에러 screenshot 넣기
					Image image = null;
					try {
						image = Image.getInstance(screenshot);
						
						image.scaleToFit(500,300);
						this.document.add(image);
						
					} catch (Exception e) {
						e.printStackTrace();
					} 
	}

	
	
	
	/* 
	 * test case 건너뛰기
	 */
	public void onTestSkipped(ITestResult result) {
		log("<<Skipped>> "+result.getName());
	}

	
	public void onFinish(ITestContext context) {
		log("END Of Test-Suite Execution->"+context.getName());
		  
		try {
			if (this.failTable != null) {
				log("Added fail table");
				P_FailedTestNum.add("Total Num: "+FailedTestNum);
				this.failTable.setSpacingBefore(15f);
				this.document.add(this.failTable);
				this.failTable.setSpacingAfter(15f);
			
			}
			
			if (this.successTable != null) {
				P_SuccessTestNum.add("Total Num: "+SuccessTestNum);
				log("Added success table");
				this.successTable.setSpacingBefore(15f);
				this.document.add(this.successTable);
				this.successTable.setSpacingBefore(15f);
			}
		} catch (DocumentException e) {
			e.printStackTrace();
		}
		
		if (this.failTable != null) {
			
			Paragraph p = new Paragraph("\nEXCEPTIONS DETAIL",
					FontFactory.getFont(FontFactory.HELVETICA, 16, Font.BOLD, new Color(255, 0, 0)));
			try {
				this.document.add(p);
			} catch (DocumentException e1) {
				e1.printStackTrace();
			}
		
			Set<Integer> keys = this.throwableMap.keySet();
			
			assert keys.size() == this.nbExceptions;
		
			for(Integer key : keys) {
				Throwable throwable = this.throwableMap.get(key);
			
				String str_throwable = throwable.getClass().toString();
				str_throwable = str_throwable.replace("class ", "");
				
				Chunk chunk = new Chunk("\n"+str_throwable,
						FontFactory.getFont(FontFactory.HELVETICA, 12, Font.BOLD, new Color(255, 0, 0)));
				chunk.setLocalDestination("" + key);
				Paragraph throwTitlePara = new Paragraph(chunk);
				
				try {
					this.document.add(throwTitlePara);
				} catch (DocumentException e3) {
					e3.printStackTrace();
				}
				
				StackTraceElement[] elems = throwable.getStackTrace();
				for(StackTraceElement ste : elems) {
					Paragraph throwParagraph = new Paragraph(ste.toString());
					try {
						this.document.add(throwParagraph);
					} catch (DocumentException e2) {
						e2.printStackTrace();
					}
				}
			}
	
	}
		this.document.close();
	}
	
	
	
	/* 
	 * console log
	 */
	public static void log(Object o) {
		System.out.println("[Listener] " + o);
	}

	

	public void onTestStart(ITestResult result) {
		// TODO Auto-generated method stub
		
	}

	public void onTestFailedButWithinSuccessPercentage(ITestResult result) {
		// TODO Auto-generated method stub
		
	}


}
