package com.skr;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.Iterator;
import javax.swing.JProgressBar;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import skr.file.csv.CsvWriter;

public class Fixer_del {

   //-- Variables
   private String postfix;
   private FileInputStream fileInputStream;
   private XSSFWorkbook readingWorkBook, writingWorkBook;
   private XSSFSheet readingSheet, writingSheet;
   private String absolutefileName;
   private java.io.File csvFile;

   private javax.swing.JFrame parentFrame;
   //private JLabel labelWorkingRecordCount, labelReadingRecordCount;
   private JProgressBar progressBarReading, progressBarWorking;

   //-- Constructors
   public Fixer_del(javax.swing.JFrame parentFrame, JProgressBar progressBarWorking, JProgressBar progressBarReading, String absoluteFileName) {

      this.postfix = "_" + System.nanoTime();
      this.absolutefileName = absoluteFileName;
      this.parentFrame = parentFrame;
      this.progressBarReading = progressBarReading;
      this.progressBarWorking = progressBarWorking;
      //this.labelWorkingRecordCount = labelWorkingRecordCount;
      //this.labelReadingRecordCount = labelReadingRecordCount;

      try {
         csvFile = createDumpCSVFile();
         System.out.println("com.skr.Fixer.ReadingThread.run(): created csv file: " + csvFile.getAbsoluteFile());
      } catch (IOException ex) {
         System.out.println("com.skr.Fixer.ReadingThread.<init>(): IOException:\n" + ex);
      }

      try {
         this.fileInputStream = new FileInputStream(absoluteFileName);

         //Create ReadingWorkbook instance holding reference to .xlsx file
         this.readingWorkBook = new XSSFWorkbook(this.fileInputStream);
         this.readingSheet = this.readingWorkBook.getSheetAt(0);

         //Create WritingWorkbook instance holding reference to .xlsx file
         this.writingWorkBook = new XSSFWorkbook();
         this.writingSheet = this.writingWorkBook.createSheet("fixed_" + this.readingSheet.getSheetName());

         //Get first/desired sheet from the workbook
         this.readingSheet = this.readingWorkBook.getSheetAt(0);
         System.out.println("com.skr.Fixer.loadWorkBookAndSheet(): reading sheet[0]: " + this.readingSheet.getSheetName());

      } catch (Exception e) {
         String msg = "Something went worng at:\ncom.skr.Fixer.loadWorkBookAndSheet()\n" + e.toString() + "\n\nPrograme will exit now.";
         javax.swing.JOptionPane.showMessageDialog(this.parentFrame, msg);
         this.close();
         System.exit(0);
         //e.printStackTrace();
      }
   }

   public FileInputStream getFileInputStream() {
      return fileInputStream;
   }

   public void setFileInputStream(FileInputStream fileInputStream) {
      this.fileInputStream = fileInputStream;
   }

   //-
   public int getRowCount() {
      return this.readingSheet.getLastRowNum();
   }

   private void close() {
      IOUtils.closeQuietly(this.fileInputStream);
   }

   private void showMSG(String msg) {
      javax.swing.JOptionPane.showMessageDialog(parentFrame, msg);
   }

   private void writeOutBook() {

      String ext = ".xlsx";

      String fileName = this.absolutefileName.substring(0, this.absolutefileName.length() - ext.length()) + postfix + ext;

      System.out.println("com.skr.Fixer.writeOutBook(): fileName= " + fileName);
      FileOutputStream out = null;

      try {
         out = new FileOutputStream(new java.io.File(fileName));

         this.writingWorkBook.write(out);
         this.showMSG("Fixing complete..\n\nFile: \"" + fileName + "\"");

         //-- ending
         close();
         System.exit(0);
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

   private java.io.File createDumpCSVFile() throws IOException {
      return File.createTempFile("duplicate-fixer_", ".csv");
   }

   //--
   private class FixingThread extends Thread {

      @Override
      public void run() {
         try {
            Iterator<Row> rowIterator = readingSheet.iterator();
            int writingRowCount = 0;

            while (rowIterator.hasNext()) {
               Utils.sleeping();
               //labelWorkingRecordCount.setText("" + writingRowCount);

               Row readingRow = rowIterator.next();
               Row writingRow = writingSheet.createRow(writingRowCount++);

               System.out.println("\n----");
               System.out.println("com.skr.Fixer.fixingThread.run() Row: " + writingRowCount);

               //For each row, iterate through all the columns
               Iterator<Cell> cellIterator = readingRow.cellIterator();
               int writingCellNum = 0;
               System.out.println("com.skr.Fixer.fixingThread.run(): readingRow.getPhysicalNumberOfCells()=" + readingRow.getPhysicalNumberOfCells());
               while (cellIterator.hasNext()) {
                  Cell readingCell = cellIterator.next();

                  //Check the cell type and format accordingly
                  System.out.println("com.skr.Fixer.fixingThread.run(): readingCell.getCellType()=" + readingCell.getCellType());

                  switch (readingCell.getCellType()) {
                     case Cell.CELL_TYPE_NUMERIC: {
                        //-- Cell of Identifier UPC/GTIN/EAN
                        System.out.println("com.skr.Fixer.fixingThread.run(): switch(readingCell.getCellType())=CELL_TYPE_NUMERIC");
                        Cell writingCell = writingRow.createCell(writingCellNum++);
                        writingCell.setCellType(Cell.CELL_TYPE_NUMERIC);
                        writingCell.setCellValue(readingCell.getNumericCellValue());
                        break;
                     }
                     case Cell.CELL_TYPE_STRING: {
                        System.out.println("com.skr.Fixer.fixingThread.run(): switch(readingCell.getCellType())=CELL_TYPE_STRING");
                        Cell writingCell = writingRow.createCell(writingCellNum++);
                        writingCell.setCellType(Cell.CELL_TYPE_STRING);
                        writingCell.setCellValue(readingCell.getStringCellValue());
                        break;
                     }
                     case Cell.CELL_TYPE_FORMULA: {
                        System.out.println("com.skr.Fixer.fixingThread.run(): switch(readingCell.getCellType())=CELL_TYPE_FORMULA");
                        Cell writingCell = writingRow.createCell(writingCellNum++);
                        writingCell.setCellType(Cell.CELL_TYPE_FORMULA);
                        writingCell.setCellValue(readingCell.getCellFormula());
                        break;
                     }
                     case Cell.CELL_TYPE_BLANK: {
                        System.out.println("com.skr.Fixer.fixingThread.run(): switch(readingCell.getCellType())=CELL_TYPE_BLANK");
                        Cell writingCell = writingRow.createCell(writingCellNum++);
                        writingCell.setCellType(Cell.CELL_TYPE_STRING);
                        writingCell.setCellValue("");
                        break;
                     }
                     case Cell.CELL_TYPE_BOOLEAN: {
                        System.out.println("com.skr.Fixer.fixingThread.run(): switch(readingCell.getCellType())=CELL_TYPE_BOOLEAN");
                        Cell writingCell = writingRow.createCell(writingCellNum++);
                        writingCell.setCellType(Cell.CELL_TYPE_BOOLEAN);
                        writingCell.setCellValue(readingCell.getBooleanCellValue());
                        break;
                     }
                     case Cell.CELL_TYPE_ERROR: {
                        System.out.println("com.skr.Fixer.fixingThread.run(): switch(readingCell.getCellType())=CELL_TYPE_ERROR");
                        Cell writingCell = writingRow.createCell(writingCellNum++);
                        writingCell.setCellType(Cell.CELL_TYPE_ERROR);
                        writingCell.setCellValue(readingCell.getErrorCellValue());
                        break;
                     }
                     default:
                        break;
                  } // End of switch
               }
            }

            writeOutBook();

         } catch (Exception e) {
            e.printStackTrace();
         }
      }
   }

   private class ReadingThread extends Thread {

      private void writeToCSV(java.util.ArrayList<String> list) {

         // before we open the file check to see if it already exists
         boolean alreadyExists = csvFile.exists();

         try {
            // use FileWriter constructor that specifies open for appending
            CsvWriter csvOutput = new CsvWriter(new FileWriter(csvFile, true), ',');

            // if the file didn't already exist then we need to write out the header line
            if (!alreadyExists) {
               for (String item : list) {
                  csvOutput.write(item);
               }
               csvOutput.endRecord();
            } else {
               // else assume that the file already has the correct header line
               for (String item : list) {
                  csvOutput.write(item);
               }
               csvOutput.endRecord();
            }
            csvOutput.close();
         } catch (Exception e) {
            e.printStackTrace();
         }

      }

      @Override
      public void run() {

         try {
            Iterator<Row> rowIterator = readingSheet.iterator();
            int readingRowCount = 0, totalRecords = getRowCount();

            while (rowIterator.hasNext()) {
               Row readingRow = rowIterator.next();
               readingRowCount++;
               java.util.ArrayList<String> list = new java.util.ArrayList<>();

               //For each row, iterate through all the columns
               Iterator<Cell> cellIterator = readingRow.cellIterator();
               while (cellIterator.hasNext()) {
                  //-- 1st Cell will be Identifier UPC/GTIN/EAN
                  Cell readingCell = cellIterator.next();

                  switch (readingCell.getCellType()) {
                     case Cell.CELL_TYPE_NUMERIC: {
                        list.add(new BigDecimal(readingCell.getNumericCellValue() + "").longValue() + "");
                        break;
                     }
                     case Cell.CELL_TYPE_STRING: {
                        list.add(readingCell.getStringCellValue() + "");
                        break;
                     }
                     case Cell.CELL_TYPE_FORMULA: {
                        list.add(readingCell.getCellFormula() + "");
                        break;
                     }
                     case Cell.CELL_TYPE_BLANK: {
                        list.add("");
                     }
                     case Cell.CELL_TYPE_BOOLEAN: {
                        list.add(readingCell.getBooleanCellValue() + "");
                        break;
                     }
                     case Cell.CELL_TYPE_ERROR: {
                        list.add(readingCell.getErrorCellValue() + "");
                        break;
                     }
                     default: {
                        break;
                     }
                  } // End of switch
               } // End of While loop

               Utils.sleeping();
               this.writeToCSV(list);
               System.out.println("com.skr.Fixer.ReadingThread.run(): writing: " + list.get(0));
               //labelReadingRecordCount.setText("" + readingRowCount);
               progressBarReading.setValue((readingRowCount * 100) / totalRecords);

            }

         } catch (Exception e) {
            showMSG("Exception while reading file:\ncom.skr.Fixer.ReadingThread.run()\n\n" + e);
         }
      }
   }

   /*
   private class updateCountThread extends Thread {

      @Override
      public void run() {
         try {
            labelWorkingRecordCount.setText("" + writingRowCount);
            System.out.println("com.skr.Fixer_del.updateCountThread.run(): writingRowCount=" + writingRowCount);
            sleep(1);
         } catch (InterruptedException ex) {
            System.out.println("com.skr.Fixer_del.updateCountThread.run(): InterruptedException:\n" + ex);
         }
      }
   }
    */
   public void doFix() {
      //new FixingThread().start();
      //new updateCountThread().start();
      new ReadingThread().start();
   }
}
