package Utilities;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
 
import org.apache.commons.*;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.streaming.SheetDataWriter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.xb.xsdschema.TotalDigitsDocument.TotalDigits;
 
public class ExcelOutputFile {
	public static FileInputStream fi;
	public static XSSFWorkbook xb;
	public static XSSFSheet ws;
	public static XSSFRow row;
	public static XSSFCell cell;
	public static void setCellData(String xlsheet,int rownum,int colnum,String data) throws IOException
	{
		String xlfile = System.getProperty("user.dir") + "\\testData\\outputfile.xlsx";
		fi = new FileInputStream(xlfile);
	    xb = new XSSFWorkbook(fi);
	    ws = xb.getSheet(xlsheet);
	    // Check if the row exists, and create one if it doesn't
	    row = ws.getRow(rownum);
	    if (row == null) {
	        row = ws.createRow(rownum);
	    }
	    // Create a cell in the row and set its value
	    cell = row.createCell(colnum);
	    cell.setCellValue(data);
	    // Write the changes back to the file
	    FileOutputStream fo = new FileOutputStream(xlfile);
	    xb.write(fo);
	    // Close all the resources to prevent memory leaks
	    xb.close();
	    fi.close();
	    fo.close();
	}
	public static void createExcel() throws IOException {
	    // Define the path to the Excel file
	    String file = System.getProperty("user.dir") + "\\testData\\outputfile.xlsx";
	    // Create a file output stream to write to the Excel file
	    FileOutputStream fo = new FileOutputStream(file);
	    // Create a new Excel workbook
	    XSSFWorkbook xb = new XSSFWorkbook();
	    // Create a new sheet named "profile"
	    XSSFSheet sheet1 = xb.createSheet("profile");
	    XSSFSheet sheet2 = xb.createSheet("Feature News");
	
	    Font font = xb.createFont();//creating
		font.setBold(true);//setting font as Bold
		CellStyle style = xb.createCellStyle();//creating cellStyle
		style.setFont(font);//Applying cell style as Bold
	    // Create the first row in the sheet
	    XSSFRow rowSheet1 = sheet1.createRow(0);
	    XSSFRow rowSheet2 = sheet2.createRow(0);
	  
	    // Create cells for "NAME" and "EMAIL"
	    XSSFCell name = rowSheet1.createCell(0);
	    XSSFCell email = rowSheet1.createCell(1);
	    
	    XSSFCell Title=rowSheet2.createCell(0);
	    XSSFCell Tooltip=rowSheet2.createCell(1);
	    XSSFCell Output=rowSheet2.createCell(2);
	    
	    name.setCellValue("NAME");
	    email.setCellValue("EMAIL");
	    name.setCellStyle(style);
	    email.setCellStyle(style);
	    
	    Title.setCellValue("TITLE");
	    Tooltip.setCellValue("TOOlTIP");
	    Output.setCellValue("OUTPUT");
	    Title.setCellStyle(style);
	    Tooltip.setCellStyle(style);
	    Output.setCellStyle(style);

	    // Write the workbook to the file output stream
	    xb.write(fo);
	    // Close the workbook and the file output stream
	    xb.close();
	    fo.close();
	}
}