package workbooklib;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class wblib 
{
	XSSFWorkbook WB;
	XSSFSheet xs;
	public wblib(String Excelpath)
	{
		try {
			File src=new File(Excelpath);
			FileInputStream FIS=new FileInputStream(src);
		    WB=new XSSFWorkbook(FIS);
			xs=WB.getSheetAt(0);
		} 
		catch (Exception e)
		{
			
		System.out.println(e.getMessage());
		
		}
	}
		
		public String getdata(int sheetnumber,int rows,int column)
		{
			xs=WB.getSheetAt(sheetnumber);
			String data=xs.getRow(rows).getCell(column).getStringCellValue();
			return data;
		}
		public int getRowCount(int sheetIndex)
		{
			int row=WB.getSheetAt(sheetIndex).getLastRowNum();
			row=row+1;
			return row;
			
		}
	}

