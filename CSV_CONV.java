import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/*
	AUTHOR 		:	VISHAL KUMAR GOURAV
	EMAIL		:	vishal.vs911@gmail.com
	DESCRIPTION	:	THIS CODE IS USED TO PARSE AN EXCEL SHEET TO .vcf file
	DATE		:	04-03-2017
	HOW TO RUN 	:
		$javac -cp "/home/user/Desktop/vcf/jxl-2.6.jar" CSV_CONV.java
		"/home/user/Desktop/vcf/jxl-2.6.jar" it is the location of jxl jar file
*/
class ParseFile{
	private String excelFile;
	private String[][] arr = null;
	
	public void setFile(String excelFile) {
		this.excelFile = excelFile;
	}
	
	public String[][] parse() throws IOException{
		File inputWorkbook = new File(excelFile);
		Workbook w;
		try{
			w = Workbook.getWorkbook(inputWorkbook);
			//Sheet wise parsing
			Sheet sheet = w.getSheet(0);
			arr = new String[sheet.getColumns()][sheet.getRows()];
			for (int j = 0; j <sheet.getColumns(); j++){
				for (int i = 0; i < sheet.getRows(); i++){
					Cell cell = sheet.getCell(j, i);
					arr[j][i] = cell.getContents();
				}
			}
		}catch (BiffException e){
			e.printStackTrace();
		}
		return arr;
	}
}

public class CSV_CONV{
	public static void main(String[] args){
		//FIRST ARGUMENT SHOUD BE THE EXCEL SHEET FILE NAME
		
	}
}
