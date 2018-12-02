package com.skr;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JProgressBar;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author SKR
 */
public class Fixer extends Thread {

   private FileInputStream fin, fout;
   private XSSFWorkbook sourceWorkbook, targetWorkbook;
   private XSSFSheet sourceSheet, targetSheet;

   private JFrame parentFrame;
   private JProgressBar progressBar;
   private javax.swing.JLabel labelStatus;
   private String barcode;

   private File srcFile, dupFile, tarFile;

   //-- Constructors
   public Fixer(JFrame parentFrame, JLabel labelStatus, JProgressBar progressBar, String barcode, File srcFile, File dupFile) {

      this.parentFrame = parentFrame;
      this.progressBar = progressBar;
      this.labelStatus = labelStatus;
      this.barcode = barcode;
      this.srcFile = srcFile;
      this.dupFile = dupFile;

      try {
         this.fin = new FileInputStream(this.srcFile);

         //Create ReadingWorkbook instance holding reference to .xlsx file
         this.sourceWorkbook = new XSSFWorkbook(this.fin);
         this.sourceSheet = this.sourceWorkbook.getSheetAt(0);

         //Create WritingWorkbook instance holding reference to .xlsx file
         this.targetWorkbook = new XSSFWorkbook();
         this.targetSheet = this.targetWorkbook.createSheet("fixed_" + this.sourceSheet.getSheetName());

         //Get first/desired sheet from the workbook
         this.sourceSheet = this.sourceWorkbook.getSheetAt(0);
         System.out.println("com.skr.Fixer.<init>(): reading sheet[0]: " + this.sourceSheet.getSheetName());

      } catch (Exception e) {
         this.showMSG("Something went worng at:\ncom.skr.Fixer.<init>()\n" + e.toString() + "\n\nPrograme will exit now.");
         System.exit(0);
      }
   }

   public int getRowCount() {
      return this.sourceSheet.getLastRowNum();
   }

   private void showMSG(String msg) {
      javax.swing.JOptionPane.showMessageDialog(parentFrame, msg);
   }

   //--
   private String getCellValueAsString(Cell cell) {
      String value = "";

      switch (cell.getCellType()) {
         case Cell.CELL_TYPE_NUMERIC: {
            value = cell.getNumericCellValue() + "";
            break;
         }
         case Cell.CELL_TYPE_STRING: {
            value = cell.getStringCellValue();
            break;
         }
         case Cell.CELL_TYPE_BOOLEAN: {
            value = cell.getBooleanCellValue() + "";
            break;
         }
         default: {
            break;
         }
      }
      return value;
   }

   private String barcodeUpdater(String value) {

      if (this.barcode.equals("UPC")) {
         while (value.length() < 12) {
            value = "0" + value;
         }
         return value;
      }

      if (this.barcode.equals("EAN")) {
         while (value.length() < 13) {
            value = "0" + value;
         }
         return value;
      }

      if (this.barcode.equals("GTIN")) {
         while (value.length() < 14) {
            value = "0" + value;
         }
         return value;
      }

      return value;
   }

   private void doCopyRow(Row srcRow, Row targetRow) {

      Iterator<Cell> cellIterator = srcRow.cellIterator();
      int writingCellNum = 0;

      while (cellIterator.hasNext()) {

         Cell readingCell = cellIterator.next();

         switch (readingCell.getCellType()) {
            case Cell.CELL_TYPE_NUMERIC: {
               Cell writingCell = targetRow.createCell(writingCellNum++);
               writingCell.setCellType(Cell.CELL_TYPE_NUMERIC);
               writingCell.setCellValue(readingCell.getNumericCellValue());
               break;
            }
            case Cell.CELL_TYPE_STRING: {
               Cell writingCell = targetRow.createCell(writingCellNum++);
               writingCell.setCellType(Cell.CELL_TYPE_STRING);
               writingCell.setCellValue(readingCell.getStringCellValue());
               break;
            }
            case Cell.CELL_TYPE_FORMULA: {
               Cell writingCell = targetRow.createCell(writingCellNum++);
               writingCell.setCellType(Cell.CELL_TYPE_FORMULA);
               writingCell.setCellValue(readingCell.getCellFormula());
               break;
            }
            case Cell.CELL_TYPE_BLANK: {
               Cell writingCell = targetRow.createCell(writingCellNum++);
               writingCell.setCellType(Cell.CELL_TYPE_STRING);
               writingCell.setCellValue("");
               break;
            }
            case Cell.CELL_TYPE_BOOLEAN: {
               Cell writingCell = targetRow.createCell(writingCellNum++);
               writingCell.setCellType(Cell.CELL_TYPE_BOOLEAN);
               writingCell.setCellValue(readingCell.getBooleanCellValue());
               break;
            }
            case Cell.CELL_TYPE_ERROR: {
               Cell writingCell = targetRow.createCell(writingCellNum++);
               writingCell.setCellType(Cell.CELL_TYPE_ERROR);
               writingCell.setCellValue(readingCell.getErrorCellValue());
               break;
            }
            default:
               break;
         } // End of switch
      }
   }

   //-- Step 1: create dump/duplicate File
   private void doCopyDataIntoDupFile() {

      progressBar.setValue(0);
      progressBar.setStringPainted(true);

      FileInputStream inStream = null;
      FileOutputStream outStream = null;

      try {
         inStream = new FileInputStream(this.srcFile);
         outStream = new FileOutputStream(this.dupFile);

         byte[] buffer = new byte[1024];
         int reading;

         long totalData = this.srcFile.length(), readData = 0;

         //-- copy the file content in bytes
         while ((reading = inStream.read(buffer)) > 0) {
            Utils.sleeping();
            readData += reading;
            outStream.write(buffer, 0, reading);
            int perc = (int) (0.5d + ((double) readData / (double) totalData) * 100);
            progressBar.setValue(perc);
            //pr.setValue((int)(readData * 100.0) / totalData + 0.5);
            //pr.setValue((int)(readData * 100.0) / totalData + 0.5);
         }

         System.out.println("com.skr.Fixer.doCopy(): File copied successfully:\n" + this.dupFile.getAbsolutePath());

      } catch (IOException ex) {
         System.out.println("com.skr.Fixer.doCopy(): IOException:\n" + ex);
      } finally {
         try {
            inStream.close();
         } catch (Exception ex) {
         }
         try {
            outStream.close();
         } catch (Exception ex) {
         }
      }
   }

   //-- Step 2: validate barcodes
   private void validateBarcodes() {

      int totalRows = this.sourceSheet.getLastRowNum();

      for (int i = 1; i <= totalRows; i++) {
         Utils.sleeping();
         progressBar.setValue((i * 100) / totalRows);

         Cell cell = this.sourceSheet.getRow(i).getCell(0);
         cell.setCellType(Cell.CELL_TYPE_STRING);
         cell.setCellValue(this.barcodeUpdater(this.getCellValueAsString(cell)));
      }

   }

//-- Step 3: remove duplicate rows
   public void doFixing() {

      int totalRows = this.sourceSheet.getLastRowNum(), targetSheetRow = 1;

      for (int i = 1; i <= totalRows; i++) {

         Utils.sleeping();
         progressBar.setValue((i * 100) / totalRows);

         boolean uniqueRow = true;
         String sourceCell = getCellValueAsString(this.sourceSheet.getRow(i).getCell(0)); // Always first cell, which is barcode.

         //For targetSheet
         for (int j = 1; j < targetSheetRow; j++) {

            String targetCell = getCellValueAsString(this.targetSheet.getRow(j).getCell(0)); // Always first cell, which is barcode.

            if (sourceCell.equalsIgnoreCase(targetCell)) {
               uniqueRow = false;
               break;
            }

         }

         if (uniqueRow) {
            this.doCopyRow(this.sourceSheet.getRow(i), this.targetSheet.createRow(targetSheetRow++));
         }
      }

   }

   @Override
   public void run() {

      //--------------------------------------------------------------------------------------------------------------------------
      this.progressBar.setValue(0);
      this.labelStatus.setText("Reading your file, please wait..");

      this.doCopyDataIntoDupFile();

      //--Creating/Copying headers------------------------------------------------------------------------------------------------
      this.doCopyRow(this.sourceSheet.getRow(0), this.targetSheet.createRow(0));

      //--------------------------------------------------------------------------------------------------------------------------
//      this.progressBar.setValue(0);
//      this.labelStatus.setText("Validating your barcodes..");
//      this.validateBarcodes();

      //--------------------------------------------------------------------------------------------------------------------------
      this.progressBar.setValue(0);
      this.labelStatus.setText("Removing duplicate lines, this may take a while..");
      this.doFixing();

      this.labelStatus.setText("Finished !!!");
      this.writeOutBook();
   }

   private void writeOutBook() {

      String ext = ".xlsx", filename = this.srcFile.getAbsolutePath(), postfix = "_fixed";

      String fileName = filename.substring(0, filename.length() - ext.length()) + postfix + ext;

      System.out.println("com.skr.Fixer.writeOutBook(): fileName= " + fileName);
      FileOutputStream out = null;

      try {
         out = new FileOutputStream(new java.io.File(fileName));

         this.targetWorkbook.write(out);
         this.dupFile.delete();
         this.showMSG("Fixing complete..\n\nFile: \"" + fileName + "\"");
      } catch (Exception ex) {
         System.out.println("com.skr.Fixer.writeOutBook(): Exception:\n" + ex);
      } finally {
         if (out != null) {
            try {
               out.close();
            } catch (IOException ex) {
            }
         }
      }
   }

}
