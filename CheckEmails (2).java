package com.emails;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import java.util.TreeMap;

import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.NoSuchProviderException;
import javax.mail.Part;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.internet.MimeBodyPart;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.opencsv.CSVReader;

public class CheckEmails {
	
	public static List<ArrayList> lstLst = new ArrayList<ArrayList>();
	
	public static void check(String host, String storeType, String user,
		      String password, String keyword, String excelDirectory,String fileDownloadDirectory, String csvFileName) 
		   {
		 List<String> lstContent = new ArrayList<String>();
		 String currentDate = "";
		      try {

		      //create properties field
		      Properties properties = new Properties();
		      properties.put("mail.pop3.host", host);
		      properties.put("mail.pop3.port", "995");
		      properties.put("mail.pop3.starttls.enable", "true");
		      Session emailSession = Session.getDefaultInstance(properties);
		      //create the POP3 store object and connect with the pop server
		      Store store = emailSession.getStore("pop3s");
		      store.connect(host, user, password);
		      //create the folder object and open it
		      Folder emailFolder = store.getFolder("INBOX");
		      emailFolder.open(Folder.READ_ONLY);
		      
		      //Search the mails for Particular date, currently it is for current date
		     /* Calendar c = Calendar.getInstance();
		   // Set calendar to the absolute beginning of the current day
			   c.set(Calendar.HOUR_OF_DAY, 0);
			   c.set(Calendar.MINUTE, 0);
			   c.set(Calendar.SECOND, 0);

		     SearchTerm[] searchCriteria  = {
		   	 new SentDateTerm(ComparisonTerm.GT, c.getTime())};
		      // retrieve the messages from the folder in an array and print it
		      //Message[] messages = emailFolder.getMessages();
		      Message[] messages = emailFolder.search(searchCriteria[0]);*/
		      //Message[] messages = emailFolder.getMessages(5000, 8070);
		      Message[] messages= emailFolder.getMessages();
		      System.out.println("messages.length---" + messages.length);
		      
		      System.out.println("==========================Program Start for converting email to excel============================");
		     /* FetchProfile fp = new FetchProfile();
		      fp.add(FetchProfile.Item.ENVELOPE);
		      fp.add(FetchProfileItem.FLAGS);
		      fp.add(FetchProfileItem.CONTENT_INFO);

		      fp.add("X-mailer");
		      emailFolder.fetch(messages, fp); */// Load the profile of the messages in 1 fetch.
		      LocalTime Starttime = LocalTime.now();
		      System.out.println(Starttime);
		      
		      
		      //This for for picking mails between 2 dates
		     /* Calendar cal = Calendar.getInstance();
			  cal.add(Calendar.MONTH, -1);
			  Date fromDate = cal.getTime();
			
			  Calendar cal1 = Calendar.getInstance();
			  cal1.add(Calendar.MONTH, 0);
			  Date toDate = cal1.getTime();
			  
		      Date toDatePlusOne = new Date(toDate.getTime() + (1000 * 60 * 60 * 24));*/
		      int messageCount = 0;
		      Calendar cal = Calendar.getInstance();
			  Date todayDate = cal.getTime();
			  SimpleDateFormat dformat = new SimpleDateFormat("ddMMyy");
			  currentDate = dformat.format(todayDate);
		      for (Message message : messages) {
		    	  //if (message.getSentDate().after(fromDate) && message.getSentDate().before(toDate)) {
		    	  String messageContent = "";
			         if(message.getSubject()!= null && message.getSubject().trim().contains(keyword)) {
			        	 messageCount ++;
				    	 if(messageCount > 10)break;
			        	 String contentType = message.getContentType();
			        	 Date messageSentDate = message.getSentDate();
			        	 SimpleDateFormat formatSendDt = new SimpleDateFormat("ddMMyyHHMMSS");
			        	 String msgSendDate = formatSendDt.format(messageSentDate);
			        	 if (contentType.contains("multipart")) {
			        		  Multipart multiPart = (Multipart) message.getContent();
			        		  for (int i = 0; i < multiPart.getCount(); i++) {
			        			    MimeBodyPart part = (MimeBodyPart) multiPart.getBodyPart(i);
			        			    if (Part.ATTACHMENT.equalsIgnoreCase(part.getDisposition())) {
			        			    	downloadAttachments(part, fileDownloadDirectory, currentDate, msgSendDate, messageCount);
			        			    }else if (part.isMimeType("text/html")) {
			    			            String html = (String) part.getContent();
			    			            messageContent = messageContent + "\n" + org.jsoup.Jsoup.parse(html).text();
			    			            
			    			        }
			        			}
			        		  	messageContent = getFormattedMailContent(messageContent);
	    			            lstContent.add(messageContent);
	    			            System.out.println("messageContent-->"+messageContent);
			        		}
			         }
		    	 // }
		      }
		      
		      System.out.println("Total mached records::"+lstContent.size());
		      LocalTime endtime = LocalTime.now();
		      System.out.println(endtime);
		      //close the store and folder objects
		      emailFolder.close(false);
		      store.close();

		      } catch (NoSuchProviderException e) {
		         e.printStackTrace();
		      } catch (MessagingException e) {
		         e.printStackTrace();
		      } catch (Exception e) {
		         e.printStackTrace();
		      }finally {
		    	  if(lstContent.size()>0) {
		    		  WriteDataToExcel(lstContent, excelDirectory, currentDate);
		    		  System.out.println("==========================Program end for converting email to excel============================");
		    	  }
		    	 
			}
		   }

		   public static void main(String[] args) {

		      String host = "outlook.office365.com";// change accordingly
		      String mailStoreType = "pop3";
		      String username = "dj@bn.com";// change accordingly
		      String password = "Spring2022!!";// change accordingly
		      //String keyword = "NOOK for iOS";
		      String keyword = "NOOK for iOS";
		      String csvFileName = "nook_diagnostic_info.csv";//That needs to be download
		      String fileDownloadDirectory = "E:\\Projects\\Outlook-Mail-Read\\download\\";
		      String excelDirectory = "E:/Projects/Outlook-Mail-Read/";
		      check(host, mailStoreType, username, password, keyword, excelDirectory,fileDownloadDirectory, csvFileName);

		   }
		   
		   public static void WriteDataToExcel(List<String> lstMailContent, String excelDirectory, String currentDate) {
			   
			        XSSFWorkbook workbook = new XSSFWorkbook();
			        XSSFSheet spreadsheet = workbook.createSheet(" Mail Data ");
			        XSSFRow row;
			        // This data needs to be written (Object[])
			        Map<String, Object[]> mailData  = new TreeMap<String, Object[]>();
			         int recordCount = 1;   
			         String [] errorDesc = null;
			         String [] errorCode = null;
			         String errorDetails = null;
			         String [] details = null;
			        for(String mailTxt : lstMailContent) {
			        	if(mailTxt !=null && mailTxt!="") {
			        	String [] arrStr = mailTxt.split("###"); 
			        	if(arrStr.length>0 ) {
			        		if(arrStr.length >1 )
				        	 errorDesc = arrStr[0].split("=");
			        		if(arrStr.length >2)
				        	 errorCode = arrStr[1].split("=");
			        		if(arrStr.length >3)
				        	errorDetails = arrStr[2].substring(16, arrStr[2].length());
			        		if(arrStr.length ==4)
				        	details = arrStr[3].split("=");
				        	
				        	mailData.put("1", new Object[] { "Error Description", "Error Code", "Error Details", "details" });
						  
						    mailData.put(String.valueOf(recordCount++), new Object[] { errorDesc[1], errorCode[1], errorDetails, details[1] });
			        	}
			          }
			        }   
			  
			        Set<String> keyid = mailData.keySet();
			  
			        int rowid = 0;
			  
			        // writing the data into the sheets...
			  
			        for (String key : keyid) {
			  
			            row = spreadsheet.createRow(rowid++);
			            Object[] objectArr = mailData.get(key);
			            int cellid = 0;
			  
			            for (Object obj : objectArr) {
			                Cell cell = row.createCell(cellid++);
			                cell.setCellValue((String)obj);
			            }
			        }
			  
			        // .xlsx is the format for Excel Sheets...
			        // writing the workbook into the file...
			       
			        FileOutputStream out;
					try {
						File directory = new File(excelDirectory);
						boolean dirFlag = directory.mkdir();
						out = new FileOutputStream(new File(excelDirectory+"/"+currentDate +".xlsx"));
						workbook.write(out);
				        out.close();
					} catch (FileNotFoundException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}catch(Exception ex) {
						
					}
			}
		   
		   public static String getFormattedMailContent(String result) {
			    String formattedResult ="";
			    String errorDesc = "";
			    String errorCode = "";
			    String errorDetails = "";
			    String details = "";
			    if(result.contains("Error Description =")) {
					int indexErrorDesc = result.indexOf("Error Description = ");
		           	result = result.substring(indexErrorDesc, result.length());
		           	if(result.contains("Error Code =")) {
			           	int indexErrorCode = result.indexOf("Error Code =");
			           	errorDesc = result.substring(0 , indexErrorCode);
			           	result = result.substring(indexErrorCode, result.length());
			           	if(result.contains("Error Details =")) {
				           	int indexErrorDetails = result.indexOf("Error Details =");
				           	errorCode = result.substring(0 , indexErrorDetails);
				           	result = result.substring(indexErrorDetails, result.length());
				           	if(result.contains("details =")) {
					           	int indexDetails = result.indexOf("details =");
					           	errorDetails = result.substring(0 , indexDetails);
					           	result = result.substring(indexDetails, result.length());
					           	if(result.contains(";")) {
						           	int indexEndDetails = result.indexOf(";");
						           	details = result.substring(0, indexEndDetails);
						           	
					           	}
				           	}else {
				           		errorDetails = result;
				           		details = "details = NA ";
				           	}
			           	}else {
			           		errorDetails = "Error Details = NA ";
			           	}
		           	}else {
		           		errorCode = "Error Code = NA ";
		           	}
			    }else {
	           		errorDesc = "Error Description = NA ";
	           	}
			    formattedResult = errorDesc + "###" + errorCode + "###" + errorDetails+ "###" + details;
	           	return formattedResult;
		   }
		   
		   public static void downloadAttachments(MimeBodyPart part,String  fileDownloadDirectory, String currentDate, String msgSendDate, int messageCount) throws IOException, MessagingException {
			   String file = part.getFileName();
			   Calendar cal = Calendar.getInstance();
		    	if(file.contains(".csv")) {
		    	   File directory = new File(fileDownloadDirectory+currentDate);
		    	   if(!directory.exists()) {
		    		   directory.mkdirs();
		    	   }
		    	   int fileNameIndex = part.getFileName().indexOf(".");
		    	   String fileName = part.getFileName().substring(0, fileNameIndex);
		    	   fileName = fileName+"-"+msgSendDate+".csv";
	               part.saveFile(fileDownloadDirectory+currentDate +"/" + File.separator +fileName);
	               //convertCSVIntoExcel(fileDownloadDirectory, currentDate, fileName);
	               convertCsvToXls(fileDownloadDirectory, currentDate, fileName, messageCount);
		    	}
			}
		   public static void convertCSVIntoExcel(String directory, String currentDate, String csvFileName ) {
				LoadOptions loadOptions = new LoadOptions(FileFormatType.CSV);
				try {
				// Creating an Workbook object with CSV file path and the loadOptions
				File excelDir = new File(directory+currentDate+"/"+"excel");
				if(!excelDir.exists()) {
					excelDir.mkdir();
				}
				File file = new File(csvFileName);
				// object
				Workbook workbook = new Workbook(directory + currentDate+ "\\"+csvFileName, loadOptions);
				workbook.save(excelDir +"/"+file.getName().replace(".csv", ".xlsx") , SaveFormat.XLSX);
				}catch(Exception ex) {
					ex.printStackTrace();
				}
			  }
		   
		   public static String convertCsvToXls(String directory, String currentDate, String csvFileName, int messageCount) {
			      CSVReader reader = null;
			      File excelDir = new File(directory+currentDate+"/"+"excel");
					if(!excelDir.exists()) {
						excelDir.mkdir();
					}
				  //File file = new File(csvFileName);
			      HSSFWorkbook workBook = new HSSFWorkbook();
			      String generatedXlsFilePath = "";
			      FileOutputStream fileOutputStream = null;
			      ArrayList<String> lstHeader = new ArrayList<>();
			  	  ArrayList<String> lstHeaderValue = new ArrayList<>();
			      try {

			          String[] nextLine;
			          reader = new CSVReader(new FileReader(directory + currentDate+ "\\"+csvFileName));

			          HSSFSheet sheet = workBook.createSheet("Excel Data");
			          //int rowNum = 1;
			         
			          Row headerRow = sheet.createRow(0);
			          Row currentRow = null;
			          
			          while((nextLine = reader.readNext()) != null) {
			        	  String rowData ="";
			        	  //System.out.println("Length-->"+nextLine.length);
			              for(int i=0; i < nextLine.length; i++) {
			            	  if(nextLine.length>1) {
				            	  if(i==0 )
				            	  {
				            		  lstHeader.add(nextLine[i].toString());
				            		 
				            	  }else if(i==1 && nextLine.length==2) {
					            		  rowData = nextLine[i];
					            		  lstHeaderValue.add(nextLine[i]);
					            	  }
					            	  else if(i==2 ) {
					            		  if(lstHeaderValue.contains(rowData)) {
					            			  lstHeaderValue.remove(rowData);
					            		  }
					            		  rowData = rowData + nextLine[i]; 
					            		  lstHeaderValue.add(rowData);
					            	  }
				            	  
			            	  }
			              }
			              
			          }
			          lstLst.add(lstHeaderValue);
			         // System.out.println("Size of Lists of List>>"+lstLst.size());
			          int i =0;
			         // if(messageCount == 1) {
				          for(String header : lstHeader) {
				        	  headerRow.createCell(i).setCellValue(header);
				    		  i++;
				          }
			         //}
				      int a = 0;   
				      for(ArrayList<String> lstHeaderVal : lstLst)  {
				    	 
					    	  int k=0;
					    	  
					    	  currentRow = sheet.createRow(a+1);
					    	  System.out.println(currentRow.getRowNum());
					          for(String headerVal : lstHeaderVal) {
					        	  //System.out.println("headerVal-->"+headerVal);
					        	  currentRow.createCell(k).setCellValue(headerVal);
					        	  //currentRow.createCell(k).setCellValue(headerVal);
					    		  k++;
					          }
				    	  //}
					          a++;
					          if(a==messageCount)break;
				      }

			          generatedXlsFilePath = directory+currentDate+"/"+"excel/" + currentDate + ".xls";
			         // logger.info("The File Is Generated At The Following Location?= " + generatedXlsFilePath);

			          fileOutputStream = new FileOutputStream(generatedXlsFilePath.trim());
			          workBook.write(fileOutputStream);
			      } catch(Exception exObj) {
			          //logger.error("Exception In convertCsvToXls() Method?=  " + exObj);
			      } finally {         
			          try {

			              /**** Closing The Excel Workbook Object ****/
			              //workBook.close();

			              /**** Closing The File-Writer Object ****/
			              fileOutputStream.close();

			              /**** Closing The CSV File-ReaderObject ****/
			              reader.close();
			          } catch (IOException ioExObj) {
			              //logger.error("Exception While Closing I/O Objects In convertCsvToXls() Method?=  " + ioExObj);          
			          }
			      }

			      return generatedXlsFilePath;
			  }   

}
