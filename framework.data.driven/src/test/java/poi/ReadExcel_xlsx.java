package poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ReadExcel_xlsx 
{
	@Test
	public void readFile() throws IOException
	{
		File file=new File("C:\\Users\\Maktum\\eclipse-workspace\\framework.data.driven\\Data\\Book1.xlsx");
		FileInputStream fis=new FileInputStream(file);
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sh=wb.getSheetAt(0);
		
		Row r=sh.getRow(0);
		int rw=sh.getLastRowNum()+1;
		int co=r.getLastCellNum();
		
		for(int i=0;i<rw;i++)
		{
			System.out.println("");
			for(int j=0;j<co;j++)
			{
				String s=sh.getRow(i).getCell(j).getStringCellValue();
				System.out.print(s+"  ");
			}
		}
		wb.close();
	}

}
