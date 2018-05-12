package poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.testng.annotations.Test;

public class ReadExcel_xls 
{
	@Test(priority=1)
	public void readMethod1() throws IOException
	{
		File file=new File("C:\\Users\\Maktum\\eclipse-workspace\\framework.data.driven\\Data\\Book1.xls");
		FileInputStream fis=new FileInputStream(file);
		HSSFWorkbook wb=new HSSFWorkbook(fis);
		HSSFSheet sh=wb.getSheetAt(0);
		Row r=sh.getRow(0);
		
		int rw=sh.getLastRowNum()+1;
		int co=r.getLastCellNum();
		
		for(int i=0;i<rw;i++)				
		{
			System.out.println("");
			for(int j=0;j<co;j++)
			{
				String s=sh.getRow(i).getCell(j).getStringCellValue();
				System.out.print(s+ "   ");
			}
		}
		wb.close();
	}
	@Test(priority=2)
	public void readMethod2() throws IOException
	{
		File file=new File("C:\\Users\\Maktum\\eclipse-workspace\\framework.data.driven\\Data\\Book1.xls");
		FileInputStream fis=new FileInputStream(file);
		HSSFWorkbook wb=new HSSFWorkbook(fis);
		HSSFSheet sh=wb.getSheetAt(0);
		
		for(Row r:sh)
		{
			System.out.println("");
			for(Cell c:r)
			{
				System.out.print(c.toString()+"  ");
			}
		}
		wb.close();
	}

}






