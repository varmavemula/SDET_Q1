package excelDataReading;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ApachePOI {
	
	 public static List<Map<String, String>> GetExcelDataAsMap(String path, String sheetName) throws IOException {
		 
		 List<Map<String, String>> dataList = new ArrayList<>();
		 
		 try(FileInputStream fis = new FileInputStream(path);
			 Workbook workbook = new XSSFWorkbook(fis))
		 {
		 
				 Sheet sheet = workbook.getSheet(sheetName); // Get first sheet
				 if (sheet == null) return dataList;
				 
				 
				 Row headerRow = sheet.getRow(0); //headerRow
				 int colsCount = headerRow.getPhysicalNumberOfCells(); //count of columns in the table.
				 
			     int rowCount = sheet.getPhysicalNumberOfRows();
			     
			     System.out.println(rowCount);
			     System.out.println(colsCount);

			     DataFormatter dataformat = new DataFormatter(); //to format all the cell data to string
			
			     for (int i = 1; i < rowCount; i++) {
			    	 
			    	 
			    	 Row currRow = sheet.getRow(i);
			    	 if(currRow == null) continue;
			    	 
			    	 Map<String, String> dataMap = new LinkedHashMap<>();
			    	 
			         for(int j=0; j<colsCount;j++) {
			        	 String key = dataformat.formatCellValue(headerRow.getCell(j));
			        	 String value = dataformat.formatCellValue(currRow.getCell(j));
			        	 dataMap.put(key, value);
			         }
			         dataList.add(dataMap);
			         }
		}
		 
		catch(Exception e)
		 {
			 e.printStackTrace();
		 }
		 return dataList;
		 
		 }
	 
	 
	 public static void main(String[] args) throws IOException {
		 List<Map<String, String>> GetDataList = GetExcelDataAsMap("Diet1.xlsx","sheet1");
		 
		 Set<String> mapSize = GetDataList.get(0).keySet();
		 System.out.println(("_").repeat(90));
		 for(String header: mapSize) {
			 System.out.print(header+((" ").repeat(17-header.length())));
			 System.out.print("|");
			
			
		 }
		 
		 for(Map<String, String> row : GetDataList) {
			 System.out.println();
			 System.out.println(("_").repeat(90));
			 for(String key: mapSize) {
				 System.out.print(row.get(key)+((" ").repeat(17-row.get(key).length())));
				 System.out.print("|");
			 }
			 
			 
		 }
	 }
	}
