import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

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
		$javac -cp "/home/user/Desktop/vcf/jxl-2.4.2.jar" CSV_CONV.java
		"/home/user/Desktop/vcf/vcf/jxl-2.6.jar" it is the location of jxl jar file
		$java java CSV_CONV <input excel file> <output file location>
	HOW TO SET CLASSPATH:	
		export CLASSPATH=/home/user/Desktop/vcf/vcf/jxl-2.4.2.jar/:$CLASSPATH
	NOTE		:
		1. THIS PROGRAM WOULD ONLY WORK FOR windows 2003-07 .xls format excel workbook
		
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
			System.out.println("Invalid excel file provided!!!");
			System.out.println("Please make sure that first argument is full qualified excel file path");
			System.out.println("Also make sure that file is in .xls format i.e. 2003 excel format");
			e.printStackTrace();
		}
		return arr;
	}
}

class VCFCreater{
	private String outputFile;
	public void setOutputFile(String outputFile){
		this.outputFile=outputFile;
	}
	public void createVCFFile(String[][] data){
		
	}
}

public class CSV_CONV{
	public static void main(String[] args) throws IOException{
		String[][] data;
		ParseFile parser;
		//FIRST ARGUMENT SHOUD BE THE EXCEL SHEET FILE NAME AND THE SECOND SHOULD BE THE
		//OUTPUT FILE LOCATION WITHOUT THE ACTUAL OUTPUT FILE NAME
		if(args.length!=2){
			System.out.println("The number of command line arguments should be exactly 2.");
			System.out.println("1. First should be the name of excel file");
			System.out.println("2. Second should be the output file location without the actual file name");
			return;		
		}
		//CHECKING FOR VALIDITY OF PATHS PROVIDED FOR 2 FILES
		if(!(new File(args[0]).exists())||(new File(args[0]).isDirectory())){
			System.out.println("Excel file provided does not exsist!!!!");
			return;
		}
		if(!(new File(args[1]).exists()&&new File(args[1]).isDirectory())){
			System.out.println("Location of output file provided does not exsist!!!!");
			return;
		}

		parser=new ParseFile();
		System.out.println("Name of the excel file is:"+args[0]);
		System.out.println("Location of the output vcf file is:"+args[1]);
		parser.setFile(args[0]);
		data=parser.parse();
		System.out.println("Number of rows in data is:"+data.length);
		System.out.println("Number of columns in data is:"+data[0].length);
		for(int i=0;i<data.length;i++){
			for(int j=0;j<data[0].length;j++){
				if(data[i][j]==""||data[i][j]==" ")
					data[i][j]="NA";
			}
		}
		/*
		System.out.println("The data is...");	
		for(int i=0;i<data.length;i++){
			for(int j=0;j<data[0].length;j++)
				System.out.print(data[i][j]+"\t");
			System.out.println(" ");
		}
		*/
	}
}
