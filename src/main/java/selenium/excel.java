package selenium;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excel {

	public static void main(String[] args) throws IOException
	{
		FileInputStream ab= new FileInputStream("C:\\Users\\Ravi\\Desktop");
		XSSFWorkbook nab=new XSSFWorkbook(ab);
		int sheet= nab.getNumberOfSheets();
		for(int i=0;i<=sheet;i++)
		{
			if(nab.getSheetName(i).equalsIgnoreCase("Sheet1"))
			{
			XSSFSheet a=nab.getSheetAt(i);
		}
		}
	}
}
