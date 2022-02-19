import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Collections;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Main {

	public static void main(String[] args) throws Exception {
		String filePath = ".\\datafile\\file1.xlsx";
		String filePath2 = ".\\datafile\\file2.xlsx";
		
		String col1IdString = "";
		String col2NameString = "";
		String col3SalaryString = "";

		// Arraylist to store data of file 1 and file 2 
		ArrayList<ExcelDataStorage> excelDataStorages = new ArrayList<>();
		
		FileInputStream inputStream = new FileInputStream(filePath);
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		//Get the title of each column
		col1IdString = sheet.getRow(0).getCell(0).getStringCellValue();
		col2NameString = sheet.getRow(0).getCell(1).getStringCellValue();
		col3SalaryString = sheet.getRow(0).getCell(2).getStringCellValue();
		
		// add file1 data into arraylist 
		addExcelData(sheet, excelDataStorages);
		inputStream.close();
		
		inputStream = new FileInputStream(filePath2);
		workbook = new XSSFWorkbook(inputStream);
		sheet = workbook.getSheetAt(0);
		
		//add file2 data into arraylist
		addExcelData(sheet, excelDataStorages);
		inputStream.close();
		
		// sort the arraylist using custom comparator here we use new ExcelDataStroage to pass comparator
		Collections.sort(excelDataStorages,new ExcelDataStorage());
		
		
		
		//create new workbook for output file
		XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
		XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet");
		XSSFRow xssfRow = xssfSheet.createRow(0);
		xssfRow.createCell(0).setCellValue(col1IdString);
		xssfRow.createCell(1).setCellValue(col2NameString);
		xssfRow.createCell(2).setCellValue(col3SalaryString);
		
		//add data from arraylist to the output sheet
		for(int i=0; i<excelDataStorages.size(); i++) {
			XSSFRow tempRow = xssfSheet.createRow(i+1);
			ExcelDataStorage tempStorage = excelDataStorages.get(i);
			tempRow.createCell(0).setCellValue((Integer)tempStorage.getId());
			tempRow.createCell(1).setCellValue(tempStorage.getName());
			tempRow.createCell(2).setCellValue((Long)tempStorage.getSalary());
		}
		
		//set the color for the first row to yellow
		XSSFCellStyle tCs = xssfWorkbook.createCellStyle();
		tCs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		tCs.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		

		xssfSheet.getRow(1).getCell(0).setCellStyle(tCs);
		xssfSheet.getRow(1).getCell(1).setCellStyle(tCs);
		xssfSheet.getRow(1).getCell(2).setCellStyle(tCs);
		
		//set the color for the last row to green
		int lastIndex = excelDataStorages.size();
		XSSFCellStyle styleLast = xssfWorkbook.createCellStyle();
		styleLast.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		styleLast.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		xssfSheet.getRow(lastIndex).getCell(0).setCellStyle(styleLast);
		xssfSheet.getRow(lastIndex).getCell(1).setCellStyle(styleLast);
		xssfSheet.getRow(lastIndex).getCell(2).setCellStyle(styleLast);
		
		
		//create output file 
		String outputPathString = ".\\datafile\\output.xlsx";
		FileOutputStream outputStream = new FileOutputStream(outputPathString);
		xssfWorkbook.write(outputStream);
		outputStream.close();
	}
	
	private static void addExcelData(XSSFSheet sheet, ArrayList<ExcelDataStorage> excelDataStorages) {
		int cols = sheet.getLastRowNum();
		int rows = sheet.getRow(1).getLastCellNum();
		for(int i=1; i<=rows; i++) {
			XSSFRow row = sheet.getRow(i);
			int idCell =(int)row.getCell(0).getNumericCellValue();
			String nameCell  = row.getCell(1).getStringCellValue();
			long salaryCell = (long)row.getCell(2).getNumericCellValue();
			ExcelDataStorage excelDataStorage = new ExcelDataStorage(idCell, nameCell, salaryCell); 
			excelDataStorages.add(excelDataStorage);
		}
	}

}
