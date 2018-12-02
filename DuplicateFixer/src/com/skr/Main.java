/**
 *
 * @author Shakir
 */
package com.skr;

import java.io.IOException;
import javax.swing.JLabel;
import javax.swing.JProgressBar;
import javax.swing.JTextField;

public class Main extends javax.swing.JFrame {

   private javax.swing.JFileChooser FileChooser;
   private final String VERSION = "version: " + this.getClass().getPackage().getImplementationVersion();
   private java.io.File sourceFile, dupFile;
   private int x0, y0;

   public Main() {
      //this.setLocation();

      initComponents();

      this.progressBar.setValue(0);
      //this.progressBarReading.setStringPainted(true);

      getComboBoxID();

      //FileChooser
      this.FileChooser = new javax.swing.JFileChooser();
      this.FileChooser.setDialogTitle("Select an excelsheet only..");
      this.FileChooser.setFileSelectionMode(javax.swing.JFileChooser.FILES_ONLY);
   }

   private boolean areFilesDone() {

      this.sourceFile = new java.io.File(this.textFieldFilePath.getText());
      if (!this.sourceFile.exists()) {
         javax.swing.JOptionPane.showMessageDialog(this, "File not availabe, please check..");
         return false;
      }

      if (!this.sourceFile.getName().endsWith(".xlsx")) {
         javax.swing.JOptionPane.showMessageDialog(this, "Please select an excel file. File extension must be..\n" + ".xlsx");
         return false;
      }

      try {
         this.dupFile = java.io.File.createTempFile("duplicate-fixer_", "_" + this.sourceFile.getName());
      } catch (IOException ex) {
         javax.swing.JOptionPane.showMessageDialog(this, "Unable to create duplicate file..");
         return false;
      }

      return true;
   }

//--Getters and setters
   public String getComboBoxID() {
      return this.comboBoxID.getSelectedItem().toString();
   }

   public JLabel getLabelFileNameOnly() {
      return labelFileNameOnly;
   }

   public JTextField getTextFieldFilePath() {
      return textFieldFilePath;
   }

   public void setTextFieldFilePath(JTextField textFieldFilePath) {
      this.textFieldFilePath = textFieldFilePath;
   }

   public JProgressBar getProgressBarReading() {
      return progressBar;
   }

   public void setProgressBarReading(JProgressBar progressBarReading) {
      this.progressBar = progressBarReading;
   }

//######################################################################################################################
   @SuppressWarnings(value = "unchecked")

   // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
   private void initComponents() {

      jLabel1 = new javax.swing.JLabel();
      jSeparator1 = new javax.swing.JSeparator();
      textFieldFilePath = new javax.swing.JTextField();
      buttonOpenFile = new javax.swing.JButton();
      jButton2 = new javax.swing.JButton();
      jLabel2 = new javax.swing.JLabel();
      buttonStartFixing = new javax.swing.JButton();
      jLabel3 = new javax.swing.JLabel();
      jSeparator2 = new javax.swing.JSeparator();
      labelFileNameOnly = new javax.swing.JLabel();
      buttonHelp = new javax.swing.JButton();
      jSeparator3 = new javax.swing.JSeparator();
      progressBar = new javax.swing.JProgressBar();
      comboBoxID = new javax.swing.JComboBox<>();
      jLabel7 = new javax.swing.JLabel();
      labelStatus = new javax.swing.JLabel();

      setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
      setTitle("< Duplicate Fixer />");
      setLocation(this.x0, this.y0);
      setMinimumSize(new java.awt.Dimension(400, 300));
      setResizable(false);

      jLabel1.setFont(new java.awt.Font("Consolas", 1, 24)); // NOI18N
      jLabel1.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
      jLabel1.setText("<DuplicateFixer/>");

      textFieldFilePath.setFont(new java.awt.Font("Tahoma", 1, 12)); // NOI18N
      textFieldFilePath.setHorizontalAlignment(javax.swing.JTextField.RIGHT);
      textFieldFilePath.setText("-- NO FILE SELECTED --");

      buttonOpenFile.setText("Open file");
      buttonOpenFile.addActionListener(new java.awt.event.ActionListener() {
         public void actionPerformed(java.awt.event.ActionEvent evt) {
            buttonOpenFileActionPerformed(evt);
         }
      });

      jButton2.setText("Force Exit");
      jButton2.addActionListener(new java.awt.event.ActionListener() {
         public void actionPerformed(java.awt.event.ActionEvent evt) {
            buttonForceExitActionPerformed(evt);
         }
      });

      jLabel2.setText("Made By: @ZULKAIFA");

      buttonStartFixing.setText("Start fixing");
      buttonStartFixing.addActionListener(new java.awt.event.ActionListener() {
         public void actionPerformed(java.awt.event.ActionEvent evt) {
            buttonStartFixingActionPerformed(evt);
         }
      });

      labelFileNameOnly.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);

      buttonHelp.setText("Help?");
      buttonHelp.addActionListener(new java.awt.event.ActionListener() {
         public void actionPerformed(java.awt.event.ActionEvent evt) {
            buttonHelpActionPerformed(evt);
         }
      });

      progressBar.setForeground(new java.awt.Color(102, 102, 255));

      comboBoxID.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
      comboBoxID.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "UPC", "EAN", "GTIN" }));

      jLabel7.setText("Vesrion 1.0");

      labelStatus.setText("Press start to fix!!");

      javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
      getContentPane().setLayout(layout);
      layout.setHorizontalGroup(
         layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
         .addComponent(jSeparator1, javax.swing.GroupLayout.Alignment.TRAILING)
         .addComponent(jSeparator2)
         .addComponent(jSeparator3)
         .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
            .addContainerGap()
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
               .addComponent(progressBar, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
               .addGroup(javax.swing.GroupLayout.Alignment.LEADING, layout.createSequentialGroup()
                  .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                     .addComponent(labelFileNameOnly, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                     .addComponent(comboBoxID, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                  .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                  .addComponent(jLabel3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                  .addGap(78, 78, 78)
                  .addComponent(buttonOpenFile, javax.swing.GroupLayout.PREFERRED_SIZE, 170, javax.swing.GroupLayout.PREFERRED_SIZE))
               .addComponent(jLabel1, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
               .addComponent(textFieldFilePath, javax.swing.GroupLayout.Alignment.LEADING)
               .addGroup(javax.swing.GroupLayout.Alignment.LEADING, layout.createSequentialGroup()
                  .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                     .addComponent(jLabel2)
                     .addComponent(jLabel7))
                  .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                  .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                     .addComponent(jButton2, javax.swing.GroupLayout.Alignment.TRAILING)
                     .addComponent(buttonHelp, javax.swing.GroupLayout.Alignment.TRAILING)))
               .addComponent(labelStatus, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
               .addGroup(layout.createSequentialGroup()
                  .addGap(0, 0, Short.MAX_VALUE)
                  .addComponent(buttonStartFixing, javax.swing.GroupLayout.PREFERRED_SIZE, 170, javax.swing.GroupLayout.PREFERRED_SIZE)))
            .addContainerGap())
      );
      layout.setVerticalGroup(
         layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
         .addGroup(layout.createSequentialGroup()
            .addContainerGap()
            .addComponent(jLabel1)
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
            .addComponent(jSeparator1, javax.swing.GroupLayout.PREFERRED_SIZE, 12, javax.swing.GroupLayout.PREFERRED_SIZE)
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
            .addComponent(textFieldFilePath, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
               .addGroup(layout.createSequentialGroup()
                  .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                     .addGroup(layout.createSequentialGroup()
                        .addGap(17, 17, 17)
                        .addComponent(jLabel3))
                     .addComponent(comboBoxID, javax.swing.GroupLayout.PREFERRED_SIZE, 27, javax.swing.GroupLayout.PREFERRED_SIZE))
                  .addGap(17, 17, 17)
                  .addComponent(jSeparator3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                  .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                  .addComponent(labelFileNameOnly))
               .addComponent(buttonOpenFile, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE))
            .addGap(18, 18, 18)
            .addComponent(labelStatus)
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addComponent(progressBar, javax.swing.GroupLayout.PREFERRED_SIZE, 19, javax.swing.GroupLayout.PREFERRED_SIZE)
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
            .addComponent(buttonStartFixing, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
            .addComponent(jSeparator2, javax.swing.GroupLayout.PREFERRED_SIZE, 4, javax.swing.GroupLayout.PREFERRED_SIZE)
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
               .addComponent(buttonHelp)
               .addComponent(jLabel2))
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
               .addComponent(jButton2)
               .addComponent(jLabel7))
            .addContainerGap())
      );

      pack();
      setLocationRelativeTo(null);
   }// </editor-fold>//GEN-END:initComponents

   private void buttonStartFixingActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_buttonStartFixingActionPerformed
      System.out.println("com.skr.Main.buttonStartFixingActionPerformed(): getSelectedItem: " + this.comboBoxID.getSelectedItem());
      System.out.println("com.skr.Main.buttonStartFixingActionPerformed(): fixing has been started");
      this.buttonStartFixing.setEnabled(false);

      if (this.areFilesDone()) {
         this.buttonStartFixing.setText("Fixing.. Please wait");
         this.comboBoxID.setEnabled(false);
         this.buttonOpenFile.setEnabled(false);

         Fixer fixer = new com.skr.Fixer(this, this.labelStatus, this.progressBar, this.comboBoxID.getSelectedItem().toString(), this.sourceFile, this.dupFile);
         fixer.start();
      } else {
         this.buttonStartFixing.setEnabled(true);
      }
   }//GEN-LAST:event_buttonStartFixingActionPerformed

   private void buttonOpenFileActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_buttonOpenFileActionPerformed
      System.out.println("com.skr.Main.OpenFileButtonActionPerformed(): Open file button clicked");
      if (javax.swing.JFileChooser.APPROVE_OPTION == this.FileChooser.showOpenDialog(this)) {
         java.io.File aFile = this.FileChooser.getSelectedFile();
         System.out.println("com.skr.Main.buttonOpenFileActionPerformed(): selected file: [" + aFile.getAbsoluteFile() + "]");
         this.textFieldFilePath.setText(aFile.getAbsolutePath());
         //this.labelFileNameOnly.setText(aFile.getName());
      }
   }//GEN-LAST:event_buttonOpenFileActionPerformed

   private void buttonForceExitActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_buttonForceExitActionPerformed
      int userVal = javax.swing.JOptionPane.showConfirmDialog(this, "Are you sure to exit??", "Force Exit?", javax.swing.JOptionPane.OK_CANCEL_OPTION, javax.swing.JOptionPane.QUESTION_MESSAGE);
      if (userVal == javax.swing.JOptionPane.OK_OPTION) {
         System.out.println("com.skr.Main.buttonForceExitActionPerformed(): exiting");
         System.gc();
         System.exit(0);
      }

   }//GEN-LAST:event_buttonForceExitActionPerformed

   private void buttonHelpActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_buttonHelpActionPerformed

      String msg = ""
              + "1) File must be an excel sheet (.xlsx)\n"
              + "2) Excel sheet can have only one sheet/tab.\n"
              + "3) First column should be valid identifier ie. UPC/EAN/GTIN\n"
              + "\nPlease contact @zulkaifa for more";
      javax.swing.JOptionPane.showMessageDialog(this, msg, "Instructions", javax.swing.JOptionPane.INFORMATION_MESSAGE);

//      java.net.URL fileUrl = this.getClass().getResource("licence.txt");
//
//      File file = new File(fileUrl.getFile());
//      System.out.println("com.skr.Main.buttonHelpActionPerformed(): " + file.getAbsolutePath());
//      if (file.exists()) {
//         try {
//            java.awt.Desktop.getDesktop().open(file);
//         } catch (java.io.IOException e) {
//            // TODO Auto-generated catch block
//            javax.swing.JOptionPane.showMessageDialog(this, e);
//            e.printStackTrace();
//         }
//      } else {
//          javax.swing.JOptionPane.showMessageDialog(this, "Something went wrong..\nFile not found");
//      }
   }//GEN-LAST:event_buttonHelpActionPerformed

   public static void main(String args[]) {
      /* Set the Nimbus look and feel */
      //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
      /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
          * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html
       */
      try {

         Thread.sleep(5000);

         for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
            //System.out.println("com.skr.Main.main(): " +info.getName());
            if ("Windows".equals(info.getName())) {
               javax.swing.UIManager.setLookAndFeel(info.getClassName());
               break;
            }
         }
      } catch (InterruptedException | ClassNotFoundException | InstantiationException | IllegalAccessException | javax.swing.UnsupportedLookAndFeelException ex) {
         System.out.println("com.skr.Main.main(): Exception:\n" + ex);
      }
      //</editor-fold>

      //</editor-fold>

      /* Create and display the form */
      java.awt.EventQueue.invokeLater(new Runnable() {
         public void run() {
            new Main().setVisible(true);
         }
      });
   }
   // Variables declaration - do not modify//GEN-BEGIN:variables
   private javax.swing.JButton buttonHelp;
   private javax.swing.JButton buttonOpenFile;
   private javax.swing.JButton buttonStartFixing;
   private javax.swing.JComboBox<String> comboBoxID;
   private javax.swing.JButton jButton2;
   private javax.swing.JLabel jLabel1;
   private javax.swing.JLabel jLabel2;
   private javax.swing.JLabel jLabel3;
   private javax.swing.JLabel jLabel7;
   private javax.swing.JSeparator jSeparator1;
   private javax.swing.JSeparator jSeparator2;
   private javax.swing.JSeparator jSeparator3;
   private javax.swing.JLabel labelFileNameOnly;
   private javax.swing.JLabel labelStatus;
   private javax.swing.JProgressBar progressBar;
   private javax.swing.JTextField textFieldFilePath;
   // End of variables declaration//GEN-END:variables
}
