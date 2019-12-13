package ddplib;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class workbook 
{
	public static void main(String[] args) throws Exception, IOException
	{
		File src=new File("D:\\Test Excel.xlsx");
		FileInputStream FIS=new FileInputStream(src);
		XSSFWorkbook WB=new XSSFWorkbook(FIS);
		XSSFSheet xs=WB.getSheetAt(0);
		int rows=xs.getLastRowNum();
		System.out.println("Number of rows"+rows);
		for(int i=0;i<rows;i++)
		{
			String value=(xs.getRow(i).getCell(i).getStringCellValue());
			System.out.println("number of columns"+value);
		}
		WB.close();
	}

}
