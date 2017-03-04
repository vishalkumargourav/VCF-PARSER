/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package vcf.converter;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.io.FileWriter;  

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
/**
 *
 * @author user
 */
class ParseFile {

    private String excelFile;
    private String[][] arr = null;

    public void setFile(String excelFile) {
        this.excelFile = excelFile;
    }

    public String[][] parse(){
        File inputWorkbook = new File(excelFile);
        Workbook w;
        try {
            w = Workbook.getWorkbook(inputWorkbook);
            //Sheet wise parsing
            Sheet sheet = w.getSheet(0);
            arr = new String[sheet.getColumns()][sheet.getRows()];
            for (int j = 0; j < sheet.getColumns(); j++) {
                for (int i = 0; i < sheet.getRows(); i++) {
                    Cell cell = sheet.getCell(j, i);
                    arr[j][i] = cell.getContents();
                }
            }
        } catch (BiffException e) {
            System.out.println("Invalid excel file provided!!!");
            System.out.println("Please make sure that first argument is full qualified excel file path");
            System.out.println("Also make sure that file is in .xls format i.e. 2003 excel format");
            e.printStackTrace();
        }catch(IOException e){
            e.printStackTrace();
        }
        return arr;
    }
}
