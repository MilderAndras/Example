/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package vezerles;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

/**
 *
 * @author András
 */
public class CikkekSzures extends Thread{
    private final MainJPanel mainJPanel1;
    private final File forgalmak;
    private final String CHAR_SET;
    private final int aktualisHonap, aktualisEv;
    private final int CIKKEK_OSZLOP = 0;
    private final int ELOZOEVIOSSZESFORGALOM_OSZLOP = 1;
    private final int ELOZOEVITORTFORGALOM_OSZLOP = 2;
    private final int IDEIOSSZESFORGALOM_OSZLOP = 3;
    private final int ELOZOEVIOSSZESARRES_OSZLOP = 4;
    private final int ELOZOEVITORTESARRES_OSZLOP = 5;
    private final int IDEIOSSZESSARRES_OSZLOP = 6;
    private final int ELSOFORGALOMHONAP_OSZLOP = 7;
    private int ELSOARRESHONAP_OSZLOP = 7;
    private final int FORRASCIKKEK_OSZLOP = 0;
    private final int FORRASEV_OSZLOP = 4;
    private final int FORRASOSSZESFORGALOM_OSZLOP = 19;
    private final int FORRASELSOFORGALOMHONAP_OSZLOP = 20;
    private final int FORRASOSSZESSARRES_OSZLOP = 32;
    private final int FORRASELSOARRESHONAP_OSZLOP = 33;

    public CikkekSzures(MainJPanel mainJPanel1, File forgalmak, String CHAR_SET, int aktualisHonap, int aktualisEv) {
        this.mainJPanel1 = mainJPanel1;
        this.forgalmak = forgalmak;
        this.CHAR_SET = CHAR_SET;
        this.aktualisHonap = aktualisHonap;
        ELSOARRESHONAP_OSZLOP += aktualisHonap;
        this.aktualisEv = aktualisEv;
    }

    @Override
    public void run() {
        HSSFWorkbook szallitoWorkbook = new HSSFWorkbook();
        HSSFSheet szallitoSheet = szallitoWorkbook.createSheet("cikkek");
        Scanner raktarScanner;
        try {
            raktarScanner = new Scanner(forgalmak, CHAR_SET);
            raktarScanner.useDelimiter("\t");
            // fejléc létrehozás
            Row fejlecRow = szallitoSheet.createRow(0);
            fejlecLetrehoz(fejlecRow);
            szures(szallitoSheet);
            fajlKiiras(szallitoWorkbook);
        } catch (FileNotFoundException ex) {
            JOptionPane.showMessageDialog(mainJPanel1, ex, "Hiba!", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void fejlecLetrehoz(Row fejlecRow) {
        int cellNum = 0;
        fejlecRow.createCell(cellNum++).setCellValue("cikkek");
        fejlecRow.createCell(cellNum++).setCellValue("előzőÉviÖsszesForgalom");
        fejlecRow.createCell(cellNum++).setCellValue("előzőÉviTörtForgalom");
        fejlecRow.createCell(cellNum++).setCellValue("ideiÖsszesForgalom");
        fejlecRow.createCell(cellNum++).setCellValue("előzőÉviÖsszesÁrrés");
        fejlecRow.createCell(cellNum++).setCellValue("előzőÉviTörtÁrrés");
        fejlecRow.createCell(cellNum++).setCellValue("ideiOsszesÁrrés");
        try {

            for (int i = 1; i <= aktualisHonap; i++) {
                fejlecRow.createCell(cellNum++).setCellValue(i + ".hónapForgalma");
            }
            for (int i = 1; i <= aktualisHonap; i++) {
                fejlecRow.createCell(cellNum++).setCellValue(i + ".hónapÁrrése");
            }

        } catch (Exception e) {
            JOptionPane.showMessageDialog(mainJPanel1, "Rossz hónap formátum, csak számot használj");
        }
    }

    private void fajlKiiras(HSSFWorkbook szallitoWorkbook) {
        JFileChooser chooser = new JFileChooser(System.getProperty("user.dir"));
        chooser.setDialogTitle("Mentés");
        chooser.setFileFilter(new FileNameExtensionFilter("xls táblázat", "xls"));
        int returnVal = chooser.showSaveDialog(mainJPanel1);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            try {
                FileOutputStream out = new FileOutputStream(new File(chooser.getSelectedFile().getPath()));
                szallitoWorkbook.write(out);
                out.close();
            } catch (FileNotFoundException ex) {
                Logger.getLogger(RaktarSzures.class.getName()).log(Level.SEVERE, null, ex);
            } catch (IOException ex) {
                Logger.getLogger(RaktarSzures.class.getName()).log(Level.SEVERE, null, ex);
            }
        }

    }

    private void szures(HSSFSheet szallitoSheet) throws FileNotFoundException {
        int rowNum = 1;
        Scanner forgalmakScanner = new Scanner(forgalmak, CHAR_SET);
        forgalmakScanner.useDelimiter("\t");
        //előolvasás
        Row kimenetRow = szallitoSheet.createRow(rowNum++);
        String[] forgalomSor = forgalmakScanner.nextLine().split("\t");
        kimenetRow.createCell(CIKKEK_OSZLOP).setCellValue(forgalomSor[FORRASCIKKEK_OSZLOP]);
        if (Integer.parseInt(forgalomSor[FORRASEV_OSZLOP]) == (aktualisEv - 1900)) {
            kimenetRow.createCell(IDEIOSSZESFORGALOM_OSZLOP).setCellValue(Double.parseDouble(forgalomSor[FORRASOSSZESFORGALOM_OSZLOP]));
            kimenetRow.createCell(IDEIOSSZESSARRES_OSZLOP).setCellValue(Double.parseDouble(forgalomSor[FORRASOSSZESSARRES_OSZLOP]));
            for (int i = 0; i < aktualisHonap; i++) {
                kimenetRow.createCell(ELSOFORGALOMHONAP_OSZLOP + i).setCellValue(Double.parseDouble(forgalomSor[FORRASELSOFORGALOMHONAP_OSZLOP + i]));
            }
            for (int i = 0; i < aktualisHonap; i++) {
                kimenetRow.createCell(ELSOARRESHONAP_OSZLOP + i).setCellValue(Double.parseDouble(forgalomSor[FORRASELSOARRESHONAP_OSZLOP + i]));
            }
        } else if(Integer.parseInt(forgalomSor[FORRASEV_OSZLOP]) == (aktualisEv - 1901)) {
            kimenetRow.createCell(ELOZOEVIOSSZESFORGALOM_OSZLOP).setCellValue(Double.parseDouble(forgalomSor[FORRASOSSZESFORGALOM_OSZLOP]));
            double elozoEviTortForgalom = 0, elozoEviTortArres = 0;
            for (int i = 0; i < aktualisHonap; i++) {
                elozoEviTortForgalom += Double.parseDouble(forgalomSor[FORRASELSOFORGALOMHONAP_OSZLOP + i]);
            }
            kimenetRow.createCell(ELOZOEVITORTFORGALOM_OSZLOP).setCellValue(elozoEviTortForgalom);
            kimenetRow.createCell(ELOZOEVIOSSZESARRES_OSZLOP).setCellValue(Double.parseDouble(forgalomSor[FORRASOSSZESSARRES_OSZLOP]));
            for (int i = 0; i < aktualisHonap; i++) {
                elozoEviTortArres += Double.parseDouble(forgalomSor[FORRASELSOARRESHONAP_OSZLOP + i]);
            }
            kimenetRow.createCell(ELOZOEVITORTESARRES_OSZLOP).setCellValue(elozoEviTortArres);
        }

        while (forgalmakScanner.hasNext()) {
            forgalomSor = forgalmakScanner.nextLine().split("\t");
            int i = 1;
            while (i < rowNum) {
                if (forgalomSor[FORRASCIKKEK_OSZLOP].equals(szallitoSheet.getRow(i).getCell(CIKKEK_OSZLOP).getStringCellValue())) {
                    kimenetRow = szallitoSheet.getRow(i);
                    if (Integer.parseInt(forgalomSor[FORRASEV_OSZLOP]) == (aktualisEv - 1900)) {
                        kimenetRow.getCell(IDEIOSSZESFORGALOM_OSZLOP,Row.CREATE_NULL_AS_BLANK).setCellValue(
                                Double.parseDouble(forgalomSor[FORRASOSSZESFORGALOM_OSZLOP])
                                + kimenetRow.getCell(IDEIOSSZESFORGALOM_OSZLOP).getNumericCellValue());
                        kimenetRow.getCell(IDEIOSSZESSARRES_OSZLOP,Row.CREATE_NULL_AS_BLANK).setCellValue(
                                Double.parseDouble(forgalomSor[FORRASOSSZESSARRES_OSZLOP])
                                + kimenetRow.getCell(IDEIOSSZESSARRES_OSZLOP).getNumericCellValue());
                        for (i = 0; i < aktualisHonap; i++) {
                            kimenetRow.getCell(ELSOFORGALOMHONAP_OSZLOP + i,Row.CREATE_NULL_AS_BLANK).setCellValue(
                                    Double.parseDouble(forgalomSor[FORRASELSOFORGALOMHONAP_OSZLOP + i])
                                    + kimenetRow.getCell(ELSOFORGALOMHONAP_OSZLOP + i).getNumericCellValue());
                        }
                        for (i = 0; i < aktualisHonap; i++) {
                            kimenetRow.getCell(ELSOARRESHONAP_OSZLOP + i,Row.CREATE_NULL_AS_BLANK).setCellValue(
                                    Double.parseDouble(forgalomSor[FORRASELSOARRESHONAP_OSZLOP + i])
                                    + kimenetRow.getCell(ELSOARRESHONAP_OSZLOP + i).getNumericCellValue());
                        }
                    } else {
                        kimenetRow.getCell(ELOZOEVIOSSZESFORGALOM_OSZLOP,Row.CREATE_NULL_AS_BLANK).setCellValue(
                                Double.parseDouble(forgalomSor[FORRASOSSZESFORGALOM_OSZLOP])
                                + kimenetRow.getCell(ELOZOEVIOSSZESFORGALOM_OSZLOP).getNumericCellValue());
                        double elozoEviTortForgalom = 0, elozoEviTortArres = 0;
                        for (i = 0; i < aktualisHonap; i++) {
                            elozoEviTortForgalom += Double.parseDouble(forgalomSor[FORRASELSOFORGALOMHONAP_OSZLOP + i]);
                        }
                        kimenetRow.getCell(ELOZOEVITORTFORGALOM_OSZLOP,Row.CREATE_NULL_AS_BLANK).setCellValue(
                                elozoEviTortForgalom + kimenetRow.getCell(ELOZOEVITORTFORGALOM_OSZLOP).getNumericCellValue());
                        kimenetRow.getCell(ELOZOEVIOSSZESARRES_OSZLOP,Row.CREATE_NULL_AS_BLANK).setCellValue(
                                Double.parseDouble(forgalomSor[FORRASOSSZESSARRES_OSZLOP])
                                + kimenetRow.getCell(ELOZOEVIOSSZESARRES_OSZLOP).getNumericCellValue());
                        for (i = 0; i < aktualisHonap; i++) {
                            elozoEviTortArres += Double.parseDouble(forgalomSor[FORRASELSOARRESHONAP_OSZLOP + i]);
                        }
                        kimenetRow.getCell(ELOZOEVITORTESARRES_OSZLOP,Row.CREATE_NULL_AS_BLANK).setCellValue(elozoEviTortArres
                                + kimenetRow.getCell(ELOZOEVITORTESARRES_OSZLOP).getNumericCellValue());
                    }
                    break;
                }
                i++;
            }
            if (i == rowNum) {
                kimenetRow = szallitoSheet.createRow(rowNum++);
                kimenetRow.createCell(CIKKEK_OSZLOP).setCellValue(forgalomSor[FORRASCIKKEK_OSZLOP]);
                if (Integer.parseInt(forgalomSor[FORRASEV_OSZLOP]) == (aktualisEv - 1900)) {
                    kimenetRow.createCell(IDEIOSSZESFORGALOM_OSZLOP).setCellValue(Double.parseDouble(forgalomSor[FORRASOSSZESFORGALOM_OSZLOP]));
                    kimenetRow.createCell(IDEIOSSZESSARRES_OSZLOP).setCellValue(Double.parseDouble(forgalomSor[FORRASOSSZESSARRES_OSZLOP]));
                    for (i = 0; i < aktualisHonap; i++) {
                        kimenetRow.createCell(ELSOFORGALOMHONAP_OSZLOP + i).setCellValue(Double.parseDouble(forgalomSor[FORRASELSOFORGALOMHONAP_OSZLOP + i]));
                    }
                    for (i = 0; i < aktualisHonap; i++) {
                        kimenetRow.createCell(ELSOARRESHONAP_OSZLOP + i).setCellValue(Double.parseDouble(forgalomSor[FORRASELSOARRESHONAP_OSZLOP + i]));
                    }
                } else {
                    kimenetRow.createCell(ELOZOEVIOSSZESFORGALOM_OSZLOP).setCellValue(Double.parseDouble(forgalomSor[FORRASOSSZESFORGALOM_OSZLOP]));
                    double elozoEviTortForgalom = 0, elozoEviTortArres = 0;
                    for (i = 0; i < aktualisHonap; i++) {
                        elozoEviTortForgalom += Double.parseDouble(forgalomSor[FORRASELSOFORGALOMHONAP_OSZLOP + i]);
                    }
                    kimenetRow.createCell(ELOZOEVITORTFORGALOM_OSZLOP).setCellValue(elozoEviTortForgalom);
                    kimenetRow.createCell(ELOZOEVIOSSZESARRES_OSZLOP).setCellValue(Double.parseDouble(forgalomSor[FORRASOSSZESSARRES_OSZLOP]));
                    for (i = 0; i < aktualisHonap; i++) {
                        elozoEviTortArres += Double.parseDouble(forgalomSor[FORRASELSOARRESHONAP_OSZLOP + i]);
                    }
                    kimenetRow.createCell(ELOZOEVITORTESARRES_OSZLOP).setCellValue(elozoEviTortArres);
                }
            }
        }
    }
}
