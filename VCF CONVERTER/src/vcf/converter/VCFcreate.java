/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package vcf.converter;

import java.io.File;
import java.io.FileWriter;

/**
 *
 * @author user
 */
public class VCFcreate {

    private String location;
    private String outputFile;
    private String name;
    private String numbers;
    private String[] no;

    public void setLocation(String location) {
        this.location = location;
    }

    public void createVCFFile(String[][] data) {
        int i;
        for (i = 0; i < 1000000; i++) {
            if (!(new File(location + "/output" + i + ".vcf").exists())) {
                break;
            }
        }
        outputFile = location + "/output" + i + ".vcf";
        System.out.println("Name of output file is:" + outputFile);
        i = 0;

        //CREATING THE OUTPUT FILE
        try {
            FileWriter fw = new FileWriter(outputFile);
            //fw.write("Test writing.");    
            for (i = 1; i < data.length; i++) {
                fw.write("BEGIN:VCARD\n");
                fw.write("VERSION:2.1\n");
                name = data[1][i];
                fw.write("N:;" + name + ";;;\n");
                fw.write("FN:" + name + "\n");
                numbers = data[4][i];
                no = numbers.replaceAll("^[,\\s]+", "").split("[,\\s]+");
                for (String n : no) {
                    if (data[6][i] != "NA") {
                        fw.write("TEL;CELL:" + data[6][i] + " " + n + "\n");
                    }
                }
                fw.write("EMAIL;HOME:" + data[5][i] + "\n");
                fw.write("ORG:" + data[2][i] + "\n");
                fw.write("TITLE:" + data[0][i] + "\n");
                fw.write("END:VCARD\n");
            }
            fw.close();
        } catch (Exception e) {
            System.out.println("Some internal error encountered!!!!UN-SUCCESSFUL");
            System.out.println(e);
        }
        System.out.println("Success...");
    }
}
