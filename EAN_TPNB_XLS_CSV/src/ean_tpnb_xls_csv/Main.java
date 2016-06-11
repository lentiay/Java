/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package ean_tpnb_xls_csv;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.lang3.StringEscapeUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author Admin
 */
public class Main {

    /**
     * @param args the command line arguments
     */

    public static void main(String[] args) {
        char BYTE_ORDER_MARK = '\uFEFF';
        try {
            String ifname = "c:\\Users\\Admin\\Documents\\NetBeansProjects\\EAN_TPNB_XLS_CSV\\ean.xlsx";
            String ofname = ifname.replaceAll(".xlsx", "");
            File fi = new File(ifname);
            FileInputStream fis = new FileInputStream(fi);
            Workbook wb = WorkbookFactory.create(fis);
            Sheet sheet = wb.getSheetAt(0);
            Row row = null;
            int fn = 1;
            FileWriter fw = new FileWriter(ofname + Integer.toString(fn) + ".csv");
            for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++){
                row = sheet.getRow(i);
                for (int j = row.getFirstCellNum(); j < row.getLastCellNum(); j++){
                    if ((i % 1000) == 0 && i != 0){
                        fw.flush();
                        fw.close();
                        fn++;
                        fw = new FileWriter(ofname + Integer.toString(fn) + ".csv");
                        fw.append(BYTE_ORDER_MARK);
                    }
                    row.getCell(j).setCellType(Cell.CELL_TYPE_STRING);
                    String str = StringEscapeUtils.escapeCsv(row.getCell(j).getStringCellValue());
                    System.out.println(StringUtils.leftPad(str, 13, '0'));
                    fw.append(StringUtils.leftPad(str, 13, '0'));
                    fw.append(";");
                    //fw.append("\n");
                }
            }
            fw.flush();
            fw.close();
            fis.close();
            wb.close();
        } catch (IOException ex) {
            Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
        } catch (InvalidFormatException ex) {
            Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
        } catch (EncryptedDocumentException ex) {
            Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

}
