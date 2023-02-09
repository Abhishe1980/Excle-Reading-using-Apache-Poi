package exceloperation;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingExcel2 {

	public static void main(String[] args) throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp Sheet");
		
	ArrayList<Object[]> empdata=new ArrayList<Object[]>();

   empdata.add(new Object[]{ "EmpId", "Name", "Job" });
   empdata.add(new Object[]{ "104", "Abhishek", "Software" });
   empdata.add(new Object[]{ "105", "Kishan", "Manager" });
   empdata.add(new Object[]{ "106", "Yogesh", "module" });
			               
	
		
		
		// Using For-Each Loop
		
		int rowCount=0;
		
		for(Object[] emp:empdata) {
			
		XSSFRow	row=sheet.createRow(rowCount++);
		int columnCount=0;
			
			for(Object value:emp)
			{
				
			XSSFCell cell=row.createCell(columnCount++);
			

			if (value instanceof String)
				cell.setCellValue((String)value);
				
				if (value instanceof Integer)
				cell.setCellValue((Integer)value);
				
				
				if (value instanceof Boolean)
				cell.setCellValue((Boolean)value);
				
			
			}
			
		}
		
		
		
		String filePath=".\\Datafile\\Countries.xlsx";
		FileOutputStream outputStream=new FileOutputStream(filePath);
				
		workbook.write(outputStream);

		outputStream.close();
		 
		System.out.println("Employe.xls is successfull written");
		
	}

}
