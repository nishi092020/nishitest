package Com.Sel.Xcel1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Xcel_Rd_Wr 
{
    public static void main( String[] args ) throws IOException, InvalidFormatException
    {
     			File f= new File("resource/Data.xlsx");
  // 			FileInputStream fis= new FileInputStream(f);//byte
  //    			XSSFWorkbook xcel= new XSSFWorkbook(fis);
    			XSSFWorkbook xcel= new XSSFWorkbook(f);
     			XSSFSheet sht=xcel.getSheet("Sheet1");
    			
  /*  			for (int i = 1; i < sht.getPhysicalNumberOfRows(); i++) {
    				String user=sht.getRow(i).getCell(0).getStringCellValue();
    				String pass= sht.getRow(i).getCell(1).getStringCellValue();

    				System.out.println(user);
    				System.out.println(pass);
    				System.out.println(".............");
    			}*/
     			String user=sht.getRow(1).getCell(0).getStringCellValue();
				String pass= sht.getRow(1).getCell(1).getStringCellValue();
				
    			sht.getRow(0).createCell(3).setCellValue("status");
    			sht.getRow(1).createCell(3).setCellValue("Fail");
	
//    			FileOutputStream fos= new FileOutputStream(f);//byte
//    			xcel.write(fos);
    			/*ERRR IN ABOVE CODE-*/
    			xcel.close();
    			//fos.close();

    }
}
