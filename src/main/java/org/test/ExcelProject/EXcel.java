package org.test.ExcelProject;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EXcel {

	public static void main(String[] args) throws IOException {
		File f=new File("C:\\Users\\RESHMA\\Desktop\\dp\\ExcelProject\\tstdata\\par.xlsx");
		FileInputStream fin=new FileInputStream(f);
		Workbook w=new XSSFWorkbook(fin);
		Sheet s = w.getSheet("Datas");
		for (int i = 0; i <s.getPhysicalNumberOfRows(); i++) {
			Row r = s.getRow(i);
			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
				Cell c = r.getCell(j);
				//System.out.println(c);
				int type = c.getCellType();
				if(type==1)
				{
					String name = c.getStringCellValue();
					System.out.println(name);
				}
				if(type==0)
				{
					double d = c.getNumericCellValue();
					long l=(long)d;
					String name = String.valueOf(l);
					System.out.println(name);
				}
			}
		}
		
		
		
		
		
		
		
	}

}
