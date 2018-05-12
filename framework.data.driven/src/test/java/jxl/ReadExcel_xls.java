package jxl;

import java.io.File;
import java.io.IOException;

import org.testng.annotations.Test;

import jxl.read.biff.BiffException;

public class ReadExcel_xls 
{
	@Test
	public void readFile() throws BiffException, IOException
	{
		File file =new File("C:\\Users\\Maktum\\eclipse-workspace\\framework.data.driven\\Data\\Book1.xls");
		Workbook wb=Workbook.getWorkbook(file);
		Sheet sh=wb.getSheet(0);
		
		int rw=sh.getRows();
		int co=sh.getColumns();
		
		for(int i=0;i<rw;i++)
		{
			System.out.println("");
			for(int j=0;j<co;j++)
			{
				String out=sh.getCell(j, i).getContents();
				System.out.print(out+ "  ");
			}
		}
	}
}
