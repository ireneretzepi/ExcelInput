/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package kitreturn;

import java.awt.Color;
import java.awt.Desktop;
import java.awt.Font;
import java.io.BufferedWriter;
import java.text.SimpleDateFormat;
import java.text.DateFormat;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Scanner;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileSystemView;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author irene
 */
public class ExcelInput extends javax.swing.JPanel {

    /**
     * Creates new form ExcelInput
     */
    public ExcelInput() {
        initComponents();
        SaveButton.setEnabled(false);
        CompleteButton1.setEnabled(false);
        CompleteButton2.setEnabled(false);
        CompleteButton3.setEnabled(false);
        CompleteButton4.setEnabled(false);
        CompleteButton5.setEnabled(false);
        CompleteButton6.setEnabled(false);
        CompleteButton7.setEnabled(false);
        CompleteButton8.setEnabled(false);
        CompleteButton9.setEnabled(false);
        ViewButton1.setEnabled(false);
        ViewButton2.setEnabled(false);
        ViewButton3.setEnabled(false);
        ViewButton4.setEnabled(false);
        ViewButton5.setEnabled(false);
        ViewButton6.setEnabled(false);
        ViewButton7.setEnabled(false);
        ViewButton8.setEnabled(false);
        ViewButton9.setEnabled(false);
        ViewButton10.setEnabled(false);
        RemoveExcelButton.setEnabled(false);
        start = true;
        
    }

    /**
     * This method is used for the buttons Add Excel File Every time a this
     * button is pressed a file chooser appears It also checks if it has an xlsx
     * extension
     */
    public void addExcel() {
        if (start==true){
        jComboBox1.setSelectedIndex(0);
        jComboBox2.setSelectedIndex(0);
        jComboBox3.setSelectedIndex(0);
        jComboBox4.setSelectedIndex(0);
        jComboBox5.setSelectedIndex(0);
        jComboBox6.setSelectedIndex(0);
        jComboBox7.setSelectedIndex(0);
        jComboBox8.setSelectedIndex(0);
        jComboBox9.setSelectedIndex(0);
        jComboBox10.setSelectedIndex(0);
        }
        start =false;
        JFileChooser chooser = new JFileChooser();
        chooser.setCurrentDirectory(new File(System.getProperty("user.home")));
        int returnValue = chooser.showOpenDialog(this);
        File selectedFile = null;
        JOptionPane optionPane = new JOptionPane();
        JOptionPane exist = new JOptionPane();
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            selectedFile = chooser.getSelectedFile();
            this.readExcel(selectedFile.getAbsolutePath());
        } else {

        }
        try {
            String filename = selectedFile.getName();
            String extension = filename.substring(filename.lastIndexOf(".") + 1, filename.length());
            String excel = "xlsx";

            while ((extension.compareTo(excel)) != 0) {
                optionPane.showMessageDialog(null, "Choose an excel file!");
                chooser.setCurrentDirectory(new File(System.getProperty("user.home")));
                returnValue = chooser.showOpenDialog(this);
                if (returnValue == JFileChooser.APPROVE_OPTION) {
                    selectedFile = chooser.getSelectedFile();
                    this.readExcel(selectedFile.getAbsolutePath());
                } else {

                }
                filename = selectedFile.getName();
                extension = filename.substring(filename.lastIndexOf(".") + 1, filename.length());
            }

            for (int i = 0; i < FileName.size(); i++) {
                while (((String) FileName.get(i)).equalsIgnoreCase(filename)) {
                    exist.showMessageDialog(null, "The File you chose already exists!");
                    chooser.setCurrentDirectory(new File(System.getProperty("user.home")));
                    returnValue = chooser.showOpenDialog(this);
                    if (returnValue == JFileChooser.APPROVE_OPTION) {
                        selectedFile = chooser.getSelectedFile();
                        this.readExcel(selectedFile.getAbsolutePath());
                    }
                    filename = selectedFile.getName();
                    extension = filename.substring(filename.lastIndexOf(".") + 1, filename.length());

                }
            }
            FilePath.add(chooser.getSelectedFile().toPath().toString());
            FileName.add(filename);
            jComboBox1.addItem(filename);
            jComboBox2.addItem(filename);
            jComboBox3.addItem(filename);
            jComboBox4.addItem(filename);
            jComboBox5.addItem(filename);
            jComboBox6.addItem(filename);
            jComboBox7.addItem(filename);
            jComboBox8.addItem(filename);
            jComboBox9.addItem(filename);
            jComboBox10.addItem(filename);
            RemoveExcelButton.setEnabled(true);

        } catch (NullPointerException e) {

        }

    }

    /*
     This method takes information that are used to restore a previous state of the application
     It saves its state in a txt file 
     */
    public void saveFile() throws IOException {
        JFileChooser chooser = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
        int returnValue = chooser.showSaveDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            BufferedWriter writer = new BufferedWriter(new FileWriter(chooser.getSelectedFile()));
            writer.write("FileName ArrayList:");
            writer.newLine();
            for (int i = 0; i < FileName.size(); i++) {
                writer.write(FileName.get(i));
                writer.newLine();
            }
            writer.write("FilePath ArrayList:");
            writer.newLine();
            for (int i = 0; i < FilePath.size(); i++) {
                writer.write(FilePath.get(i));
                writer.newLine();
            }
            ArrayList<String> SelectedComboBoxes = new ArrayList<>();
            if ((jComboBox1.getSelectedIndex() != 0) && (jComboBox1.getSelectedIndex() != -1)) {
                SelectedComboBoxes.add(jComboBox1.getSelectedItem().toString());
            }
            if ((jComboBox2.getSelectedIndex() != 0) && (jComboBox2.getSelectedIndex() != -1)) {
                SelectedComboBoxes.add(jComboBox2.getSelectedItem().toString());
            }
            if ((jComboBox3.getSelectedIndex() != 0) && (jComboBox3.getSelectedIndex() != -1)) {
                SelectedComboBoxes.add(jComboBox3.getSelectedItem().toString());
            }
            if ((jComboBox4.getSelectedIndex() != 0) && (jComboBox4.getSelectedIndex() != -1)) {
                SelectedComboBoxes.add(jComboBox4.getSelectedItem().toString());
            }
            if ((jComboBox5.getSelectedIndex() != 0) && (jComboBox5.getSelectedIndex() != -1)) {
                SelectedComboBoxes.add(jComboBox5.getSelectedItem().toString());
            }
            if ((jComboBox6.getSelectedIndex() != 0) && (jComboBox6.getSelectedIndex() != -1)) {
                SelectedComboBoxes.add(jComboBox6.getSelectedItem().toString());
            }
            if ((jComboBox7.getSelectedIndex() != 0) && (jComboBox7.getSelectedIndex() != -1)) {
                SelectedComboBoxes.add(jComboBox7.getSelectedItem().toString());
            }
            if ((jComboBox8.getSelectedIndex() != 0) && (jComboBox8.getSelectedIndex() != -1)) {
                SelectedComboBoxes.add(jComboBox8.getSelectedItem().toString());
            }
            if ((jComboBox9.getSelectedIndex() != 0) && (jComboBox9.getSelectedIndex() != -1)) {
                SelectedComboBoxes.add(jComboBox9.getSelectedItem().toString());
            }
            if ((jComboBox10.getSelectedIndex() != 0) && (jComboBox10.getSelectedIndex() != -1)) {
                SelectedComboBoxes.add(jComboBox10.getSelectedItem().toString());
            }
            writer.write("Selected Combo Boxes:");
            writer.newLine();
            for (int i = 0; i < SelectedComboBoxes.size(); i++) {
                writer.write(SelectedComboBoxes.get(i));
                writer.newLine();
            }
            writer.write("Number of up/down buttons that should be disabled:" + completeButtonsPressed.size());
            writer.newLine();
            writer.write("View files created:");
            writer.newLine();
            for (int i = 0; i < viewFiles.size(); i++) {
                writer.write(viewFiles.get(i));
                writer.newLine();
            }
            writer.write("Complete Buttons Pressed:");
            writer.newLine();
            for (int i = 0; i < completeButtonsPressed.size(); i++) {
                writer.write(completeButtonsPressed.get(i));
                writer.newLine();
            }
            writer.close();
        }
    }

    /*
    Reads the txt file with the information of the previous state of the GUI and updates it the current
    to match the previous
     */
    public void readFile(ArrayList<String> FileName1, ArrayList<String> FilePath1, ArrayList<String> SelectedComboBoxes1, String disableUpDown, ArrayList<String> ViewFiles1, ArrayList<String> completeButtonsPressed1) {

        for (int i = 0; i < FileName1.size(); i++) {
            FileName.add(FileName1.get(i));
            System.out.println("FileName1.get(i) " + FileName1.get(i));
            String name = FileName1.get(i);
            jComboBox1.addItem(name);
            jComboBox2.addItem(name);
            jComboBox3.addItem(name);
            jComboBox4.addItem(name);
            jComboBox5.addItem(name);
            jComboBox6.addItem(name);
            jComboBox7.addItem(name);
            jComboBox8.addItem(name);
            jComboBox9.addItem(name);
            jComboBox10.addItem(name);
            FilePath.add(FilePath1.get(i));
        }
        for (int j = 0; j < SelectedComboBoxes1.size(); j++) {
            OpenSelectedComboBoxes.add(SelectedComboBoxes1.get(j));
        }
        ArrayList<Integer> indexOfSelectedItem = new ArrayList<>();
        for (int j = 0; j < OpenSelectedComboBoxes.size(); j++) {
            for (int p = 0; p < FileName.size(); p++) {
                if (FileName.get(p).equalsIgnoreCase(OpenSelectedComboBoxes.get(j))) {
                    indexOfSelectedItem.add(p + 1);
                }
            }
        }
        for (int i = 0; i < completeButtonsPressed1.size(); i++) {
            if (completeButtonsPressed1.get(i).equals("CompleteButton1")) {
                Complete1();
                ViewButton1.setEnabled(true);
                try {
                    // TODO add your handling code here:
                    Runtime.getRuntime().exec("excel " + ViewFiles1.get(i));
                } catch (IOException ex) {
                    Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
                }

            }
            if (completeButtonsPressed1.get(i).equals("CompleteButton2")) {
                complete2();
                ViewButton2.setEnabled(true);
                try {
                    // TODO add your handling code here:
                    Runtime.getRuntime().exec("excel " + ViewFiles1.get(i));
                } catch (IOException ex) {
                    Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
                }
            } else if (completeButtonsPressed1.get(i).equals("CompleteButton3")) {
                complete3();
                ViewButton3.setEnabled(true);
                try {
                    // TODO add your handling code here:
                    Runtime.getRuntime().exec("excel " + ViewFiles1.get(i));
                } catch (IOException ex) {
                    Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
                }
            } else if (completeButtonsPressed1.get(i).equals("CompleteButton4")) {
                complete4();
                ViewButton4.setEnabled(true);
                try {
                    // TODO add your handling code here:
                    Runtime.getRuntime().exec("excel " + ViewFiles1.get(i));
                } catch (IOException ex) {
                    Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
                }
            } else if (completeButtonsPressed1.get(i).equals("CompleteButton5")) {
                complete5();
                ViewButton5.setEnabled(true);
                try {
                    // TODO add your handling code here:
                    Runtime.getRuntime().exec("excel " + ViewFiles1.get(i));
                } catch (IOException ex) {
                    Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
                }
            } else if (completeButtonsPressed1.get(i).equals("CompleteButton6")) {
                complete6();
                ViewButton6.setEnabled(true);
                try {
                    // TODO add your handling code here:
                    Runtime.getRuntime().exec("excel " + ViewFiles1.get(i));
                } catch (IOException ex) {
                    Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
                }
            } else if (completeButtonsPressed1.get(i).equals("CompleteButton7")) {
                complete7();
                ViewButton7.setEnabled(true);
                try {
                    // TODO add your handling code here:
                    Runtime.getRuntime().exec("excel " + ViewFiles1.get(i));
                } catch (IOException ex) {
                    Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
                }
            } else if (completeButtonsPressed1.get(i).equals("CompleteButton8")) {
                complete8();
                ViewButton8.setEnabled(true);
                try {
                    // TODO add your handling code here:
                    Runtime.getRuntime().exec("excel " + ViewFiles1.get(i));
                } catch (IOException ex) {
                    Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
                }
            } else if (completeButtonsPressed1.get(i).equals("CompleteButton9")) {
                complete9();
                ViewButton9.setEnabled(true);
                try {
                    // TODO add your handling code here:
                    Runtime.getRuntime().exec("excel " + ViewFiles1.get(i));
                } catch (IOException ex) {
                    Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
                }
                ViewButton10.setEnabled(true);
                /*
                try {
                    // TODO add your handling code here:
                    Runtime.getRuntime().exec("excel " + ViewFiles1.get(i + 1));
                } catch (IOException ex) {
                    Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
                }
                */

            }
        }

        
        System.out.println("OpenSelectedComboBoxes " + OpenSelectedComboBoxes);
        for (int b = 0; b < ViewFiles1.size(); b++) {
            viewFiles.add(ViewFiles1.get(b));
        }
        completeButtonsPressed = completeButtonsPressed1;
        
        System.out.println("indexOfSelectedItem " + indexOfSelectedItem);
        for (int k = 0; k < indexOfSelectedItem.size(); k++) {
            if (k == 0) {
                jComboBox1.setSelectedIndex(indexOfSelectedItem.get(0));
            }
            if (k == 1) {
                jComboBox2.setSelectedIndex(indexOfSelectedItem.get(1));
            }
            if (k == 2) {
                jComboBox3.setSelectedIndex(indexOfSelectedItem.get(2));
            }
            if (k == 3) {
                jComboBox4.setSelectedIndex(indexOfSelectedItem.get(3));
            }
            if (k == 4) {
                jComboBox5.setSelectedIndex(indexOfSelectedItem.get(4));
            }
            if (k == 5) {
                jComboBox6.setSelectedIndex(indexOfSelectedItem.get(5));
            }
            if (k == 6) {
                jComboBox7.setSelectedIndex(indexOfSelectedItem.get(6));
            }
            if (k == 7) {
                jComboBox8.setSelectedIndex(indexOfSelectedItem.get(7));
            }
            if (k == 8) {
                jComboBox9.setSelectedIndex(indexOfSelectedItem.get(8));
            }
            if (k == 9) {
                jComboBox10.setSelectedIndex(indexOfSelectedItem.get(9));
            }
        }

    }

    public ArrayList getFilePath() {
        return FilePath;
    }

    public void readExcel(String testFileName) {

    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jMenuBar1 = new javax.swing.JMenuBar();
        jMenu1 = new javax.swing.JMenu();
        jMenu2 = new javax.swing.JMenu();
        jMenuBar2 = new javax.swing.JMenuBar();
        jMenu3 = new javax.swing.JMenu();
        jMenu4 = new javax.swing.JMenu();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        jLabel6 = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        jLabel8 = new javax.swing.JLabel();
        jLabel9 = new javax.swing.JLabel();
        jLabel10 = new javax.swing.JLabel();
        jLabel11 = new javax.swing.JLabel();
        AddExcelButton = new javax.swing.JButton();
        jComboBox1 = new javax.swing.JComboBox<>();
        RemoveExcelButton = new javax.swing.JButton();
        jComboBox2 = new javax.swing.JComboBox<>();
        downButton1 = new javax.swing.JButton();
        CompleteButton1 = new javax.swing.JButton();
        ViewButton1 = new javax.swing.JButton();
        jComboBox3 = new javax.swing.JComboBox<>();
        jComboBox4 = new javax.swing.JComboBox<>();
        jComboBox5 = new javax.swing.JComboBox<>();
        jComboBox6 = new javax.swing.JComboBox<>();
        jComboBox7 = new javax.swing.JComboBox<>();
        jComboBox8 = new javax.swing.JComboBox<>();
        jComboBox9 = new javax.swing.JComboBox<>();
        jComboBox10 = new javax.swing.JComboBox<>();
        upButton1 = new javax.swing.JButton();
        downButton2 = new javax.swing.JButton();
        CompleteButton2 = new javax.swing.JButton();
        ViewButton2 = new javax.swing.JButton();
        upButton2 = new javax.swing.JButton();
        downButton3 = new javax.swing.JButton();
        CompleteButton3 = new javax.swing.JButton();
        ViewButton3 = new javax.swing.JButton();
        upButton3 = new javax.swing.JButton();
        upButton4 = new javax.swing.JButton();
        upButton5 = new javax.swing.JButton();
        upButton6 = new javax.swing.JButton();
        upButton7 = new javax.swing.JButton();
        upButton8 = new javax.swing.JButton();
        upButton9 = new javax.swing.JButton();
        downButton4 = new javax.swing.JButton();
        downButton5 = new javax.swing.JButton();
        downButton6 = new javax.swing.JButton();
        downButton7 = new javax.swing.JButton();
        downButton8 = new javax.swing.JButton();
        downButton9 = new javax.swing.JButton();
        CompleteButton4 = new javax.swing.JButton();
        CompleteButton5 = new javax.swing.JButton();
        CompleteButton6 = new javax.swing.JButton();
        CompleteButton7 = new javax.swing.JButton();
        CompleteButton8 = new javax.swing.JButton();
        CompleteButton9 = new javax.swing.JButton();
        ViewButton4 = new javax.swing.JButton();
        ViewButton5 = new javax.swing.JButton();
        ViewButton6 = new javax.swing.JButton();
        ViewButton7 = new javax.swing.JButton();
        ViewButton8 = new javax.swing.JButton();
        ViewButton9 = new javax.swing.JButton();
        ViewButton10 = new javax.swing.JButton();
        SaveButton = new javax.swing.JButton();

        jMenu1.setText("File");
        jMenuBar1.add(jMenu1);

        jMenu2.setText("Edit");
        jMenuBar1.add(jMenu2);

        jMenu3.setText("File");
        jMenuBar2.add(jMenu3);

        jMenu4.setText("Edit");
        jMenuBar2.add(jMenu4);

        setOpaque(false);
        setPreferredSize(new java.awt.Dimension(1000, 800));

        jLabel1.setText("Input the Excel files in the order the projects will be completed!");

        jLabel2.setText("1.");

        jLabel3.setText("2.");

        jLabel4.setText("3.");

        jLabel5.setText("4.");

        jLabel6.setText("5.");

        jLabel7.setText("6.");

        jLabel8.setText("7.");

        jLabel9.setText("8.");

        jLabel10.setText("9.");

        jLabel11.setText("10.");

        AddExcelButton.setText("Add Excel File");
        AddExcelButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                AddExcelButtonActionPerformed(evt);
            }
        });

        jComboBox1.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Select Excel File" }));
        jComboBox1.setOpaque(false);
        jComboBox1.addItemListener(new java.awt.event.ItemListener() {
            public void itemStateChanged(java.awt.event.ItemEvent evt) {
                jComboBox1ItemStateChanged(evt);
            }
        });
        jComboBox1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jComboBox1ActionPerformed(evt);
            }
        });

        RemoveExcelButton.setText("Remove Excel File");
        RemoveExcelButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                RemoveExcelButtonActionPerformed(evt);
            }
        });

        jComboBox2.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Select Excel File" }));
        jComboBox2.setOpaque(false);
        jComboBox2.addItemListener(new java.awt.event.ItemListener() {
            public void itemStateChanged(java.awt.event.ItemEvent evt) {
                jComboBox2ItemStateChanged(evt);
            }
        });

        downButton1.setText("↓");
        downButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                downButton1ActionPerformed(evt);
            }
        });

        CompleteButton1.setText("Complete");
        CompleteButton1.addChangeListener(new javax.swing.event.ChangeListener() {
            public void stateChanged(javax.swing.event.ChangeEvent evt) {
                CompleteButton1StateChanged(evt);
            }
        });
        CompleteButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                CompleteButton1ActionPerformed(evt);
            }
        });
        CompleteButton1.addKeyListener(new java.awt.event.KeyAdapter() {
            public void keyPressed(java.awt.event.KeyEvent evt) {
                CompleteButton1KeyPressed(evt);
            }
        });

        ViewButton1.setText("View👁");
        ViewButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ViewButton1ActionPerformed(evt);
            }
        });

        jComboBox3.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Select Excel File" }));
        jComboBox3.setOpaque(false);
        jComboBox3.addItemListener(new java.awt.event.ItemListener() {
            public void itemStateChanged(java.awt.event.ItemEvent evt) {
                jComboBox3ItemStateChanged(evt);
            }
        });
        jComboBox3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jComboBox3ActionPerformed(evt);
            }
        });

        jComboBox4.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Select Excel File" }));
        jComboBox4.setOpaque(false);
        jComboBox4.addItemListener(new java.awt.event.ItemListener() {
            public void itemStateChanged(java.awt.event.ItemEvent evt) {
                jComboBox4ItemStateChanged(evt);
            }
        });

        jComboBox5.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Select Excel File" }));
        jComboBox5.setOpaque(false);
        jComboBox5.addItemListener(new java.awt.event.ItemListener() {
            public void itemStateChanged(java.awt.event.ItemEvent evt) {
                jComboBox5ItemStateChanged(evt);
            }
        });
        jComboBox5.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jComboBox5ActionPerformed(evt);
            }
        });

        jComboBox6.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Select Excel File" }));
        jComboBox6.setOpaque(false);
        jComboBox6.addItemListener(new java.awt.event.ItemListener() {
            public void itemStateChanged(java.awt.event.ItemEvent evt) {
                jComboBox6ItemStateChanged(evt);
            }
        });

        jComboBox7.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Select Excel File" }));
        jComboBox7.setOpaque(false);
        jComboBox7.addItemListener(new java.awt.event.ItemListener() {
            public void itemStateChanged(java.awt.event.ItemEvent evt) {
                jComboBox7ItemStateChanged(evt);
            }
        });

        jComboBox8.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Select Excel File" }));
        jComboBox8.setOpaque(false);
        jComboBox8.addItemListener(new java.awt.event.ItemListener() {
            public void itemStateChanged(java.awt.event.ItemEvent evt) {
                jComboBox8ItemStateChanged(evt);
            }
        });

        jComboBox9.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Select Excel File" }));
        jComboBox9.setOpaque(false);
        jComboBox9.addItemListener(new java.awt.event.ItemListener() {
            public void itemStateChanged(java.awt.event.ItemEvent evt) {
                jComboBox9ItemStateChanged(evt);
            }
        });
        jComboBox9.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jComboBox9ActionPerformed(evt);
            }
        });

        jComboBox10.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Select Excel File" }));
        jComboBox10.setOpaque(false);
        jComboBox10.addItemListener(new java.awt.event.ItemListener() {
            public void itemStateChanged(java.awt.event.ItemEvent evt) {
                jComboBox10ItemStateChanged(evt);
            }
        });
        jComboBox10.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jComboBox10ActionPerformed(evt);
            }
        });

        upButton1.setText("↑");
        upButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                upButton1ActionPerformed(evt);
            }
        });

        downButton2.setText("↓");
        downButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                downButton2ActionPerformed(evt);
            }
        });

        CompleteButton2.setText("Complete");
        CompleteButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                CompleteButton2ActionPerformed(evt);
            }
        });
        CompleteButton2.addKeyListener(new java.awt.event.KeyAdapter() {
            public void keyPressed(java.awt.event.KeyEvent evt) {
                CompleteButton2KeyPressed(evt);
            }
        });

        ViewButton2.setText("View👁");
        ViewButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ViewButton2ActionPerformed(evt);
            }
        });

        upButton2.setText("↑");
        upButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                upButton2ActionPerformed(evt);
            }
        });

        downButton3.setText("↓");
        downButton3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                downButton3ActionPerformed(evt);
            }
        });

        CompleteButton3.setText("Complete");
        CompleteButton3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                CompleteButton3ActionPerformed(evt);
            }
        });

        ViewButton3.setText("View👁");
        ViewButton3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ViewButton3ActionPerformed(evt);
            }
        });

        upButton3.setText("↑");
        upButton3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                upButton3ActionPerformed(evt);
            }
        });

        upButton4.setText("↑");
        upButton4.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                upButton4ActionPerformed(evt);
            }
        });

        upButton5.setText("↑");
        upButton5.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                upButton5ActionPerformed(evt);
            }
        });

        upButton6.setText("↑");
        upButton6.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                upButton6ActionPerformed(evt);
            }
        });

        upButton7.setText("↑");
        upButton7.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                upButton7ActionPerformed(evt);
            }
        });

        upButton8.setText("↑");
        upButton8.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                upButton8ActionPerformed(evt);
            }
        });

        upButton9.setText("↑");
        upButton9.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                upButton9ActionPerformed(evt);
            }
        });

        downButton4.setText("↓");
        downButton4.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                downButton4ActionPerformed(evt);
            }
        });

        downButton5.setText("↓");
        downButton5.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                downButton5ActionPerformed(evt);
            }
        });

        downButton6.setText("↓");
        downButton6.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                downButton6ActionPerformed(evt);
            }
        });

        downButton7.setText("↓");
        downButton7.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                downButton7ActionPerformed(evt);
            }
        });

        downButton8.setText("↓");
        downButton8.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                downButton8ActionPerformed(evt);
            }
        });

        downButton9.setText("↓");
        downButton9.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                downButton9ActionPerformed(evt);
            }
        });

        CompleteButton4.setText("Complete");
        CompleteButton4.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                CompleteButton4ActionPerformed(evt);
            }
        });

        CompleteButton5.setText("Complete");
        CompleteButton5.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                CompleteButton5ActionPerformed(evt);
            }
        });

        CompleteButton6.setText("Complete");
        CompleteButton6.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                CompleteButton6ActionPerformed(evt);
            }
        });

        CompleteButton7.setText("Complete");
        CompleteButton7.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                CompleteButton7ActionPerformed(evt);
            }
        });

        CompleteButton8.setText("Complete");
        CompleteButton8.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                CompleteButton8ActionPerformed(evt);
            }
        });

        CompleteButton9.setText("Complete");
        CompleteButton9.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                CompleteButton9ActionPerformed(evt);
            }
        });

        ViewButton4.setText("View👁");
        ViewButton4.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ViewButton4ActionPerformed(evt);
            }
        });

        ViewButton5.setText("View👁");
        ViewButton5.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ViewButton5ActionPerformed(evt);
            }
        });

        ViewButton6.setText("View👁");
        ViewButton6.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ViewButton6ActionPerformed(evt);
            }
        });

        ViewButton7.setText("View👁");
        ViewButton7.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ViewButton7ActionPerformed(evt);
            }
        });

        ViewButton8.setText("View👁");
        ViewButton8.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ViewButton8ActionPerformed(evt);
            }
        });

        ViewButton9.setText("View👁");
        ViewButton9.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ViewButton9ActionPerformed(evt);
            }
        });

        ViewButton10.setText("View👁");
        ViewButton10.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ViewButton10ActionPerformed(evt);
            }
        });

        SaveButton.setText("Save");
        SaveButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                SaveButtonActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(this);
        this.setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 341, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(AddExcelButton, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(29, 29, 29)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(jLabel10)
                                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                        .addComponent(jLabel4, javax.swing.GroupLayout.Alignment.TRAILING)
                                        .addComponent(jLabel3, javax.swing.GroupLayout.Alignment.TRAILING)
                                        .addComponent(jLabel2, javax.swing.GroupLayout.Alignment.TRAILING)
                                        .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                            .addComponent(jLabel7, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                            .addComponent(jLabel6, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                            .addComponent(jLabel5, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                            .addComponent(jLabel9, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                            .addComponent(jLabel8)))
                                    .addComponent(jLabel11))
                                .addGap(28, 28, 28)
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                    .addComponent(jComboBox9, 0, 293, Short.MAX_VALUE)
                                    .addComponent(jComboBox8, 0, 293, Short.MAX_VALUE)
                                    .addComponent(jComboBox7, 0, 293, Short.MAX_VALUE)
                                    .addComponent(jComboBox6, 0, 293, Short.MAX_VALUE)
                                    .addComponent(jComboBox3, 0, 293, Short.MAX_VALUE)
                                    .addComponent(jComboBox2, 0, 293, Short.MAX_VALUE)
                                    .addComponent(jComboBox4, 0, 293, Short.MAX_VALUE)
                                    .addComponent(jComboBox1, 0, 293, Short.MAX_VALUE)
                                    .addComponent(jComboBox5, 0, 293, Short.MAX_VALUE)
                                    .addComponent(jComboBox10, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
                            .addComponent(SaveButton, javax.swing.GroupLayout.PREFERRED_SIZE, 127, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(14, 14, 14)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(layout.createSequentialGroup()
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                                            .addComponent(upButton2, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                            .addComponent(upButton1, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(downButton2)
                                            .addComponent(downButton3)))
                                    .addGroup(layout.createSequentialGroup()
                                        .addComponent(upButton3)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                        .addComponent(downButton4))
                                    .addGroup(layout.createSequentialGroup()
                                        .addComponent(upButton4)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                        .addComponent(downButton5))
                                    .addComponent(downButton1, javax.swing.GroupLayout.Alignment.TRAILING))
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(layout.createSequentialGroup()
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(RemoveExcelButton, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                        .addGap(47, 47, 47))
                                    .addGroup(layout.createSequentialGroup()
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                                            .addGroup(javax.swing.GroupLayout.Alignment.LEADING, layout.createSequentialGroup()
                                                .addGap(16, 16, 16)
                                                .addComponent(CompleteButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 124, javax.swing.GroupLayout.PREFERRED_SIZE))
                                            .addGroup(layout.createSequentialGroup()
                                                .addGap(18, 18, 18)
                                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                                    .addComponent(CompleteButton5, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                                    .addComponent(CompleteButton4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                                    .addComponent(CompleteButton3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                                    .addComponent(CompleteButton2, javax.swing.GroupLayout.DEFAULT_SIZE, 122, Short.MAX_VALUE))))
                                        .addGap(18, 18, 18)
                                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                            .addComponent(ViewButton1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                            .addComponent(ViewButton2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                            .addComponent(ViewButton3, javax.swing.GroupLayout.DEFAULT_SIZE, 175, Short.MAX_VALUE)
                                            .addComponent(ViewButton4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                            .addComponent(ViewButton5, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))))
                            .addGroup(layout.createSequentialGroup()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                    .addComponent(upButton5, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    .addComponent(upButton6, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    .addComponent(upButton7, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    .addComponent(upButton8, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    .addComponent(upButton9, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                                .addGap(10, 10, 10)
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                    .addGroup(layout.createSequentialGroup()
                                        .addComponent(downButton7)
                                        .addGap(18, 18, 18)
                                        .addComponent(CompleteButton7, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                                    .addGroup(layout.createSequentialGroup()
                                        .addComponent(downButton8)
                                        .addGap(18, 18, 18)
                                        .addComponent(CompleteButton8, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                                    .addGroup(layout.createSequentialGroup()
                                        .addComponent(downButton6)
                                        .addGap(18, 18, 18)
                                        .addComponent(CompleteButton6, javax.swing.GroupLayout.PREFERRED_SIZE, 122, javax.swing.GroupLayout.PREFERRED_SIZE))
                                    .addGroup(layout.createSequentialGroup()
                                        .addComponent(downButton9)
                                        .addGap(18, 18, 18)
                                        .addComponent(CompleteButton9, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
                                .addGap(18, 18, 18)
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(ViewButton6, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    .addComponent(ViewButton7, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    .addComponent(ViewButton8, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    .addComponent(ViewButton9, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    .addComponent(ViewButton10, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))))))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 24, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(AddExcelButton, javax.swing.GroupLayout.PREFERRED_SIZE, 46, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                        .addComponent(CompleteButton1)
                        .addComponent(ViewButton1)
                        .addComponent(downButton1))
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel2)
                            .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                    .addComponent(jLabel3)
                                    .addComponent(upButton1)
                                    .addComponent(downButton2)
                                    .addComponent(CompleteButton2)
                                    .addComponent(ViewButton2))
                                .addGap(16, 16, 16))
                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                                .addComponent(jComboBox2, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)))
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                .addComponent(jLabel4)
                                .addComponent(upButton2)
                                .addComponent(downButton3)
                                .addComponent(CompleteButton3)
                                .addComponent(ViewButton3))
                            .addComponent(jComboBox3, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jComboBox4, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(upButton3)
                            .addComponent(downButton4)
                            .addComponent(CompleteButton4)
                            .addComponent(ViewButton4)
                            .addComponent(jLabel5))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jComboBox5, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel6)
                            .addComponent(upButton4)
                            .addComponent(downButton5)
                            .addComponent(CompleteButton5)
                            .addComponent(ViewButton5))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                                .addComponent(jLabel7)
                                .addComponent(upButton5)
                                .addComponent(downButton6)
                                .addComponent(CompleteButton6)
                                .addComponent(ViewButton6))
                            .addComponent(jComboBox6, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jComboBox7, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel8)
                            .addComponent(upButton6)
                            .addComponent(downButton7)
                            .addComponent(CompleteButton7)
                            .addComponent(ViewButton7))
                        .addGap(8, 8, 8)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel9)
                            .addComponent(jComboBox8, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(upButton7)
                            .addComponent(downButton8)
                            .addComponent(CompleteButton8)
                            .addComponent(ViewButton8))))
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(6, 6, 6)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jComboBox9, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(upButton8)
                            .addComponent(downButton9)
                            .addComponent(CompleteButton9)
                            .addComponent(ViewButton9)))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(18, 18, 18)
                        .addComponent(jLabel10)))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jComboBox10, javax.swing.GroupLayout.PREFERRED_SIZE, 35, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel11)
                    .addComponent(upButton9)
                    .addComponent(ViewButton10))
                .addGap(18, 18, 18)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(0, 17, Short.MAX_VALUE)
                        .addComponent(RemoveExcelButton, javax.swing.GroupLayout.PREFERRED_SIZE, 50, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addComponent(SaveButton, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addGap(72, 72, 72))
        );
    }// </editor-fold>//GEN-END:initComponents

    private void AddExcelButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_AddExcelButtonActionPerformed
        // TODO add your handling code here:
        addExcel();
    }//GEN-LAST:event_AddExcelButtonActionPerformed

    private void jComboBox1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jComboBox1ActionPerformed
        // TODO add your handling code here:


    }//GEN-LAST:event_jComboBox1ActionPerformed

    private void RemoveExcelButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_RemoveExcelButtonActionPerformed
        // TODO add your handling code here:
        //  FilePath FileName
        //FileName covert to array
        String[] FN = new String[FileName.size()];
        for (int i = 0; i < FileName.size(); i++) {
            String add = FileName.get(i);
            FN[i] = add;
        }
        String input;
        input = (String) JOptionPane.showInputDialog(null, "Choose remove file name ",
                "Choose the name of the excel file you want to remove:",
                JOptionPane.QUESTION_MESSAGE, null,
                FN,
                FN[0]);
        if (input != null) {

            int index1 = -1;
            boolean flag = false;
            while (flag == false) {
                for (int i = 0; i < FileName.size(); i++) {
                    if ((FileName.get(i).equalsIgnoreCase(input))) {
                        index1 = i;
                        flag = true;
                    }
                }
            }
            if (index1 != -1) {
                FileName.remove(index1);
                FilePath.remove(index1);
                jComboBox1.removeItemAt(index1 + 1);
                jComboBox2.removeItemAt(index1 + 1);
                jComboBox3.removeItemAt(index1 + 1);
                jComboBox4.removeItemAt(index1 + 1);
                jComboBox5.removeItemAt(index1 + 1);
                jComboBox6.removeItemAt(index1 + 1);
                jComboBox7.removeItemAt(index1 + 1);
                jComboBox8.removeItemAt(index1 + 1);
                jComboBox9.removeItemAt(index1 + 1);
                jComboBox10.removeItemAt(index1 + 1);

            }
            if (FileName.isEmpty()) {
                RemoveExcelButton.setEnabled(false);
            }
        }
    }//GEN-LAST:event_RemoveExcelButtonActionPerformed

    private void jComboBox5ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jComboBox5ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jComboBox5ActionPerformed

    private void jComboBox10ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jComboBox10ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jComboBox10ActionPerformed

    private void downButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_downButton2ActionPerformed
        // TODO add your handling code here:
        int i = jComboBox2.getSelectedIndex();
        int j = jComboBox3.getSelectedIndex();
        jComboBox3.setSelectedIndex(i);
        jComboBox2.setSelectedIndex(j);

    }//GEN-LAST:event_downButton2ActionPerformed

    private void jComboBox9ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jComboBox9ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jComboBox9ActionPerformed

    private void downButton7ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_downButton7ActionPerformed
        // TODO add your handling code here:
        int i = jComboBox7.getSelectedIndex();
        int j = jComboBox8.getSelectedIndex();
        jComboBox8.setSelectedIndex(i);
        jComboBox7.setSelectedIndex(j);
    }//GEN-LAST:event_downButton7ActionPerformed

    private void complete4() {
        complete4 = true;
        completeButtonsPressed.add("CompleteButton4");
        if ((jComboBox5.getSelectedIndex() != 0) && (jComboBox6.getSelectedIndex() != 0)) {
            CompleteButton4.setEnabled(false);
            CompleteButton5.setEnabled(true);
        } else {
            CompleteButton4.setEnabled(false);
        }
        jComboBox4.setBackground(Color.red);
        jComboBox4.setEnabled(false);
        upButton4.setEnabled(false);
        downButton4.setEnabled(false);
        ViewButton4.setEnabled(true);
        currentExcelIndex = 4;
        int SelectedIndex = 0;
        String FileItem = (String) jComboBox4.getSelectedItem();
        int FP = 0;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(jComboBox4.getSelectedItem().toString())) {
                FP = i;
            }
        }
        System.out.println(FilePath.get(FP));
        String NameOfFile = FilePath.get(FP) + time.toString() + ".xlsx";
        viewFiles.add(NameOfFile);
        for (int i = 0; i < FileName.size(); i++) {
            if (FileItem.equalsIgnoreCase(FileName.get(i))) {
                SelectedIndex = i;
            }
        }
        addKitReturn(SelectedIndex);
    }

    private void CompleteButton4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_CompleteButton4ActionPerformed
        // TODO add your handling code here:
        complete4();
    }//GEN-LAST:event_CompleteButton4ActionPerformed

    private void ViewButton4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ViewButton4ActionPerformed
        // TODO add your handling code here:
        int FP = 0;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(jComboBox4.getSelectedItem().toString())) {
                FP = i;
            }
        }
        System.out.println(FilePath.get(FP));
        String NameOfFile = FilePath.get(FP) + time.toString() + ".xlsx";
        try {
            // TODO add your handling code here:
            Runtime.getRuntime().exec("excel " + NameOfFile);
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_ViewButton4ActionPerformed

    private void ViewButton5ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ViewButton5ActionPerformed
        // TODO add your handling code here:
        int FP = 0;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(jComboBox5.getSelectedItem().toString())) {
                FP = i;
            }
        }
        String NameOfFile = FilePath.get(FP) + time.toString() + ".xlsx";
        try {
            // TODO add your handling code here:
            Runtime.getRuntime().exec("excel " + NameOfFile);
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_ViewButton5ActionPerformed

    private void Complete1() {
        pressedComplete1 = true;
        completeButtonsPressed.add("CompleteButton1");
        if ((jComboBox2.getSelectedIndex() != 0) && (jComboBox3.getSelectedIndex() != 0)) {
            CompleteButton1.setEnabled(false);
            CompleteButton2.setEnabled(true);
        } else {
            CompleteButton1.setEnabled(false);
        }

        jComboBox1.setBackground(Color.red);
        jComboBox1.setEnabled(false);
        downButton1.setEnabled(false);
        ViewButton1.setEnabled(true);
        upButton1.setEnabled(false);
        RemoveExcelButton.setEnabled(true);
        SaveButton.setEnabled(true);
        //current project completed
        currentExcelIndex = 1;
        int SelectedIndex = 0;
        String FileItem = (String) jComboBox1.getSelectedItem();
        int FP = 0;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(jComboBox1.getSelectedItem().toString())) {
                FP = i;
            }
        }
        System.out.println(FilePath.get(FP));
        String NameOfFile = FilePath.get(FP) + time.toString() + ".xlsx";
        viewFiles.add(NameOfFile);

        for (int i = 0; i < FileName.size(); i++) {
            if (FileItem.equalsIgnoreCase(FileName.get(i))) {
                SelectedIndex = i;
            }
        }
        addKitReturn(SelectedIndex);
    }

    private void CompleteButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_CompleteButton1ActionPerformed
        // TODO add your handling code here:
        Complete1();

    }//GEN-LAST:event_CompleteButton1ActionPerformed

    public void addKitReturn(int SelectedIndex) {

        SelectedFilesPath.clear();
        currentMPN.clear();
        nextMPN.clear();
        indexofcommon.clear();
        System.out.println("selected file path after emptying it" + SelectedFilesPath);
        //the index of item in the filepath of the selected item in the combobox

        createPathArraylist();
        currentFilePath = FilePath.get(SelectedIndex);
        System.out.println("currentFilePath in complete 4" + currentFilePath);

        try {
            FileInputStream excelFile = excelFile = new FileInputStream(new File(FilePath.get(SelectedIndex)));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.rowIterator();
            DataFormatter dataFormatter = new DataFormatter();
            XSSFFont font = (XSSFFont) workbook.createFont();
            font.setBold(true);
            font.setUnderline(FontUnderline.SINGLE);
            XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
            style.setFont(font);
            boolean found = false;

            for (Row row : sheet) {
                Cell cell = row.getCell(4);
                String cellValue = dataFormatter.formatCellValue(cell);

                if (found == true) {
                    if ((cellValue.compareTo("")) != 0) {
                        currentMPN.add(cellValue);
                    }
                }
                if (cellValue.equalsIgnoreCase("Manufacturer Part Number")) {
                    found = true;
                }

            }

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                row.createCell(6);
                // For each row, iterate through all the columns 
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    String cellValue = dataFormatter.formatCellValue(cell);

                    if ((cellValue.equalsIgnoreCase("Populate"))) { //&&((sheet.getRow(6).getCell(6)))==null)
                        cell = cellIterator.next();
                        cell.setCellValue("Re-Kit");
                        cell.setCellStyle(style);
                    }

                }
            }

            excelFile.close();
            String NameOfFile = FilePath.get(SelectedIndex) + java.time.LocalDate.now() + ".xlsx";
            FileOutputStream output_file = new FileOutputStream(new File(NameOfFile));
            //write changes
            workbook.write(output_file);

        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }

        for (int i = 0; i < SelectedFilesPath.size(); i++) {
            compare(i);

        }
    }

    public void compare(int index) {
        // the file name of the file that is compared with the current one
        nextMPN.clear();
        String nextFileName = "";
        for (int i = 0; i < FilePath.size(); i++) {
            if ((FilePath.get(i).compareTo(SelectedFilesPath.get(index))) == 0) {
                nextFileName = FileName.get(i);
            }
        }
        System.out.println(nextFileName);

        try {
            System.out.println("SelectedFilesPathin index " + SelectedFilesPath.get(index));
            FileInputStream excelFile2 = new FileInputStream(new File(SelectedFilesPath.get(index)));
            Workbook workbook2 = new XSSFWorkbook(excelFile2);
            Sheet sheet2 = workbook2.getSheetAt(0);
            DataFormatter dataFormatter2 = new DataFormatter();

            boolean found2 = false;

            for (Row row2 : sheet2) {
                Cell cell2 = row2.getCell(4);
                String cellValue2 = dataFormatter2.formatCellValue(cell2);

                if (found2 == true) {
                    if ((cellValue2.compareTo("")) != 0) {
                        nextMPN.add(cellValue2);
                    }
                }
                if (cellValue2.equalsIgnoreCase("Manufacturer Part Number")) {
                    found2 = true;
                }

            }

            excelFile2.close();
            FileOutputStream output_file = new FileOutputStream(new File(SelectedFilesPath.get(index)));
            //write changes
            workbook2.write(output_file);

        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }

        ArrayList<String> current = new ArrayList<String>(currentMPN);
        ArrayList<String> common = new ArrayList<String>(currentMPN);
        current.removeAll(nextMPN);
        common.removeAll(current);
        System.out.println("common" + common);
        System.out.println("currentFilePath " + currentFilePath);
        try {
            String NameOfFile = currentFilePath + java.time.LocalDate.now() + ".xlsx";
            FileInputStream excelFile3 = excelFile3 = new FileInputStream(new File(NameOfFile));
            Workbook workbook3 = new XSSFWorkbook(excelFile3);
            Sheet sheet3 = workbook3.getSheetAt(0);
            DataFormatter dataFormatter3 = new DataFormatter();

            boolean found = false;

            int i = 0;
            for (Row row3 : sheet3) {
                Cell cell3 = row3.getCell(4);
                String cellValue3 = dataFormatter3.formatCellValue(cell3);

                if ((found == true) && (i < common.size())) {
                    if (((cellValue3 == null)) || (((common.get(i)).compareTo(cellValue3)) == 0)) {
                        indexofcommon.add(row3.getRowNum());
                        Row row4 = sheet3.getRow(row3.getRowNum());
                        Cell column = row4.getCell(6);
                        String previous = column.getStringCellValue();
                        row4.createCell(6 + index);
                        column = row4.getCell(6 + index);
                        column.setCellValue(nextFileName);
                        i = i + 1;
                    }

                }
                if (cellValue3.equalsIgnoreCase("Manufacturer Part Number")) {
                    found = true;
                }

            }
            System.out.println(indexofcommon);

            excelFile3.close();
            FileOutputStream output_file = new FileOutputStream(NameOfFile);
            //write changes
            workbook3.write(output_file);

        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

    public void createPathArraylist() {
        SelectedFilesPath.clear();
        int SelectedIndex2 = 0;
        int SelectedIndex3 = 0;
        int SelectedIndex4 = 0;
        int SelectedIndex5 = 0;
        int SelectedIndex6 = 0;
        int SelectedIndex7 = 0;
        int SelectedIndex8 = 0;
        int SelectedIndex9 = 0;
        int SelectedIndex10 = 0;

        String FileItem2 = (String) jComboBox2.getSelectedItem();
        String FileItem3 = (String) jComboBox3.getSelectedItem();
        String FileItem4 = (String) jComboBox4.getSelectedItem();
        String FileItem5 = (String) jComboBox5.getSelectedItem();
        String FileItem6 = (String) jComboBox6.getSelectedItem();
        String FileItem7 = (String) jComboBox7.getSelectedItem();
        String FileItem8 = (String) jComboBox8.getSelectedItem();
        String FileItem9 = (String) jComboBox9.getSelectedItem();
        String FileItem10 = (String) jComboBox10.getSelectedItem();

        //SelectedFilesPath the arraylist containing the paths of the selected files 
        if (complete2 == false) {
            for (int i = 0; i < FileName.size(); i++) {
                if (FileItem2.equalsIgnoreCase(FileName.get(i))) {
                    SelectedIndex2 = i;
                }
            }
            SelectedFilesPath.add(FilePath.get(SelectedIndex2));
        }

        if (complete3 == false) {
            if ((FileItem3.compareTo("Select Excel File")) != 0) {
                for (int i = 0; i < FileName.size(); i++) {
                    if (FileItem3.equalsIgnoreCase(FileName.get(i))) {
                        SelectedIndex3 = i;
                    }
                }
                SelectedFilesPath.add(FilePath.get(SelectedIndex3));
            }
        }

        if (complete4 == false) {
            if ((FileItem4.compareTo("Select Excel File")) != 0) {
                for (int i = 0; i < FileName.size(); i++) {
                    if (FileItem4.equalsIgnoreCase(FileName.get(i))) {
                        SelectedIndex4 = i;
                    }
                }
                SelectedFilesPath.add(FilePath.get(SelectedIndex4));
            }
        }

        if (complete5 == false) {
            if ((FileItem5.compareTo("Select Excel File")) != 0) {
                for (int i = 0; i < FileName.size(); i++) {
                    if (FileItem5.equalsIgnoreCase(FileName.get(i))) {
                        SelectedIndex5 = i;
                    }
                }
                SelectedFilesPath.add(FilePath.get(SelectedIndex5));
            }
        }

        if (complete6 == false) {
            if ((FileItem6.compareTo("Select Excel File")) != 0) {
                for (int i = 0; i < FileName.size(); i++) {
                    if (FileItem6.equalsIgnoreCase(FileName.get(i))) {
                        SelectedIndex6 = i;
                    }
                }
                SelectedFilesPath.add(FilePath.get(SelectedIndex6));
            }
        }

        if (complete7 == false) {
            if ((FileItem7.compareTo("Select Excel File")) != 0) {
                for (int i = 0; i < FileName.size(); i++) {
                    if (FileItem7.equalsIgnoreCase(FileName.get(i))) {
                        SelectedIndex7 = i;
                    }
                }
                SelectedFilesPath.add(FilePath.get(SelectedIndex7));
            }
        }

        if (complete8 == false) {
            if ((FileItem8.compareTo("Select Excel File")) != 0) {
                for (int i = 0; i < FileName.size(); i++) {
                    if (FileItem8.equalsIgnoreCase(FileName.get(i))) {
                        SelectedIndex8 = i;
                    }
                }
                SelectedFilesPath.add(FilePath.get(SelectedIndex8));
            }
        }

        if (complete9 == false) {
            if ((FileItem9.compareTo("Select Excel File")) != 0) {
                for (int i = 0; i < FileName.size(); i++) {
                    if (FileItem9.equalsIgnoreCase(FileName.get(i))) {
                        SelectedIndex9 = i;
                    }
                }
                SelectedFilesPath.add(FilePath.get(SelectedIndex9));
            }
        }

        if (complete10 == false) {
            if ((FileItem10.compareTo("Select Excel File")) != 0) {
                for (int i = 0; i < FileName.size(); i++) {
                    if (FileItem10.equalsIgnoreCase(FileName.get(i))) {
                        SelectedIndex10 = i;
                    }
                }
                SelectedFilesPath.add(FilePath.get(SelectedIndex10));
            }
        }
    }
    private void jComboBox1ItemStateChanged(java.awt.event.ItemEvent evt) {//GEN-FIRST:event_jComboBox1ItemStateChanged
        // TODO add your handling code here:
        if ((jComboBox1.getSelectedIndex() == 0) || (jComboBox2.getSelectedIndex() == 0) && (pressedComplete1 == false)) {
            CompleteButton1.setEnabled(false);
        }
        if ((jComboBox1.getSelectedIndex() != 0) && (jComboBox2.getSelectedIndex() != 0)) {
            CompleteButton1.setEnabled(true);

        }
    }//GEN-LAST:event_jComboBox1ItemStateChanged

    private void jComboBox2ItemStateChanged(java.awt.event.ItemEvent evt) {//GEN-FIRST:event_jComboBox2ItemStateChanged
        // TODO add your handling code here:  
        if ((jComboBox1.getSelectedIndex() != 0) && (jComboBox2.getSelectedIndex() != 0) && (pressedComplete1 == false)) {
            CompleteButton1.setEnabled(true);
        }
        if ((jComboBox1.getSelectedIndex() == 0) || (jComboBox2.getSelectedIndex() == 0) && (pressedComplete1 == false)) {
            CompleteButton1.setEnabled(false);
        }
        if ((pressedComplete1 == true) && (jComboBox2.getSelectedIndex() != 0) && (jComboBox3.getSelectedIndex() != 0)) {
            CompleteButton2.setEnabled(true);
        }
        if ((pressedComplete1 == true) && (jComboBox2.getSelectedIndex() == 0) || (jComboBox3.getSelectedIndex() == 0)) {
            CompleteButton2.setEnabled(false);
        }
        if (jComboBox2.getSelectedIndex() != 0) {
            selected2 = true;
        } else {
            selected2 = false;
        }


    }//GEN-LAST:event_jComboBox2ItemStateChanged

    private void jComboBox3ItemStateChanged(java.awt.event.ItemEvent evt) {//GEN-FIRST:event_jComboBox3ItemStateChanged
        // TODO add your handling code here:
        if ((pressedComplete1 == true) && (complete2 == false) && ((jComboBox3.getSelectedIndex() == 0) || (jComboBox2.getSelectedIndex() == 0))) {
            CompleteButton2.setEnabled(false);
        }
        if ((pressedComplete1 == true) && (complete2 == false) && (jComboBox3.getSelectedIndex() != 0) && (jComboBox2.getSelectedIndex() != 0)) {
            CompleteButton2.setEnabled(true);
        } else {
        }
        if ((pressedComplete1 == true) && (complete2 == true) && (jComboBox3.getSelectedIndex() != 0) && (jComboBox4.getSelectedIndex() != 0)) {
            CompleteButton3.setEnabled(true);
        }
        if ((pressedComplete1 == true) && (complete2 == true) && (jComboBox3.getSelectedIndex() == 0) || (jComboBox4.getSelectedIndex() == 0)) {
            CompleteButton3.setEnabled(false);
        }
        if (jComboBox3.getSelectedIndex() != 0) {
            selected3 = true;
            selected3 = false;
        }

    }//GEN-LAST:event_jComboBox3ItemStateChanged

    private void jComboBox4ItemStateChanged(java.awt.event.ItemEvent evt) {//GEN-FIRST:event_jComboBox4ItemStateChanged
        // TODO add your handling code here:
        if ((complete2 == true) && (complete3 == false) && ((jComboBox4.getSelectedIndex() == 0) || (jComboBox3.getSelectedIndex() == 0))) {
            CompleteButton3.setEnabled(false);
        }
        if ((complete2 == true) && (complete3 == false) && (jComboBox4.getSelectedIndex() != 0) && (jComboBox3.getSelectedIndex() != 0)) {
            CompleteButton3.setEnabled(true);
        }
        if ((complete2 == true) && (complete3 == true) && (jComboBox4.getSelectedIndex() != 0) && (jComboBox5.getSelectedIndex() != 0)) {
            CompleteButton4.setEnabled(true);
        }
        if ((complete2 == true) && (complete3 == true) && (jComboBox4.getSelectedIndex() == 0) || (jComboBox5.getSelectedIndex() == 0)) {
            CompleteButton4.setEnabled(false);
        }
        if (jComboBox4.getSelectedIndex() != 0) {
            selected4 = true;
        } else {
            selected4 = false;
        }
    }//GEN-LAST:event_jComboBox4ItemStateChanged

    private void jComboBox5ItemStateChanged(java.awt.event.ItemEvent evt) {//GEN-FIRST:event_jComboBox5ItemStateChanged
        // TODO add your handling code here:
        if ((complete3 == true) && (complete4 == false) && ((jComboBox5.getSelectedIndex() == 0) || (jComboBox4.getSelectedIndex() == 0))) {
            CompleteButton4.setEnabled(false);
        }
        if ((complete3 == true) && (complete4 == false) && (jComboBox5.getSelectedIndex() != 0) && (jComboBox4.getSelectedIndex() != 0)) {
            CompleteButton4.setEnabled(true);
        }
        if ((complete3 == true) && (complete4 == true) && (jComboBox5.getSelectedIndex() != 0) && (jComboBox6.getSelectedIndex() != 0)) {
            CompleteButton5.setEnabled(true);
        }
        if ((complete3 == true) && (complete4 == true) && (jComboBox5.getSelectedIndex() == 0) || (jComboBox6.getSelectedIndex() == 0)) {
            CompleteButton5.setEnabled(false);
        }
        if (jComboBox5.getSelectedIndex() != 0) {
            selected5 = true;
        } else {
            selected5 = false;
        }
    }//GEN-LAST:event_jComboBox5ItemStateChanged

    private void jComboBox6ItemStateChanged(java.awt.event.ItemEvent evt) {//GEN-FIRST:event_jComboBox6ItemStateChanged
        // TODO add your handling code here:
        if ((complete4 == true) && (complete5 == false) && ((jComboBox6.getSelectedIndex() == 0) || (jComboBox5.getSelectedIndex() == 0))) {
            CompleteButton5.setEnabled(false);
        }
        if ((complete4 == true) && (complete5 == false) && (jComboBox6.getSelectedIndex() != 0) && (jComboBox5.getSelectedIndex() != 0)) {
            CompleteButton5.setEnabled(true);
        }
        if ((complete4 == true) && (complete5 == true) && (jComboBox6.getSelectedIndex() != 0) && (jComboBox7.getSelectedIndex() != 0)) {
            CompleteButton6.setEnabled(true);
        }
        if ((complete4 == true) && (complete5 == true) && (jComboBox6.getSelectedIndex() == 0) || (jComboBox7.getSelectedIndex() == 0)) {
            CompleteButton6.setEnabled(false);
        }
        if (jComboBox6.getSelectedIndex() != 0) {
            selected6 = true;
        } else {
            selected6 = false;
        }
    }//GEN-LAST:event_jComboBox6ItemStateChanged

    private void jComboBox7ItemStateChanged(java.awt.event.ItemEvent evt) {//GEN-FIRST:event_jComboBox7ItemStateChanged
        // TODO add your handling code here:
        if ((complete5 == true) && (complete6 == false) && ((jComboBox7.getSelectedIndex() == 0) || (jComboBox6.getSelectedIndex() == 0))) {
            CompleteButton6.setEnabled(false);
        }
        if ((complete5 == true) && (complete6 == false) && (jComboBox7.getSelectedIndex() != 0) && (jComboBox6.getSelectedIndex() != 0)) {
            CompleteButton6.setEnabled(true);
        }
        if ((complete5 == true) && (complete6 == true) && (jComboBox7.getSelectedIndex() != 0) && (jComboBox8.getSelectedIndex() != 0)) {
            CompleteButton7.setEnabled(true);
        }
        if ((complete5 == true) && (complete6 == true) && (jComboBox7.getSelectedIndex() == 0) || (jComboBox8.getSelectedIndex() == 0)) {
            CompleteButton7.setEnabled(false);
        }
        if (jComboBox7.getSelectedIndex() != 0) {
            selected7 = true;
        } else {
            selected7 = false;
        }
    }//GEN-LAST:event_jComboBox7ItemStateChanged

    private void jComboBox8ItemStateChanged(java.awt.event.ItemEvent evt) {//GEN-FIRST:event_jComboBox8ItemStateChanged
        // TODO add your handling code here:
        if ((complete6 == true) && (complete7 == false) && ((jComboBox8.getSelectedIndex() == 0) || (jComboBox7.getSelectedIndex() == 0))) {
            CompleteButton7.setEnabled(false);
        }
        if ((complete6 == true) && (complete7 == false) && (jComboBox8.getSelectedIndex() != 0) && (jComboBox7.getSelectedIndex() != 0)) {
            CompleteButton7.setEnabled(true);
        }
        if ((complete6 == true) && (complete7 == true) && (jComboBox8.getSelectedIndex() != 0) && (jComboBox9.getSelectedIndex() != 0)) {
            CompleteButton8.setEnabled(true);
        }
        if ((complete6 == true) && (complete7 == true) && (jComboBox8.getSelectedIndex() == 0) || (jComboBox9.getSelectedIndex() == 0)) {
            CompleteButton8.setEnabled(false);
        }
        if (jComboBox8.getSelectedIndex() != 0) {
            selected8 = true;
        } else {
            selected8 = false;
        }
    }//GEN-LAST:event_jComboBox8ItemStateChanged

    private void jComboBox9ItemStateChanged(java.awt.event.ItemEvent evt) {//GEN-FIRST:event_jComboBox9ItemStateChanged
        // TODO add your handling code here:
        if ((complete7 == true) && (complete8 == false) && ((jComboBox9.getSelectedIndex() == 0) || (jComboBox8.getSelectedIndex() == 0))) {
            CompleteButton8.setEnabled(false);
        }
        if ((complete7 == true) && (complete8 == false) && (jComboBox9.getSelectedIndex() != 0) && (jComboBox8.getSelectedIndex() != 0)) {
            CompleteButton8.setEnabled(true);
        }
        if ((complete7 == true) && (complete8 == true) && (jComboBox9.getSelectedIndex() != 0) && (jComboBox10.getSelectedIndex() != 0)) {
            CompleteButton9.setEnabled(true);
        }
        if ((complete7 == true) && (complete8 == true) && (jComboBox9.getSelectedIndex() == 0) || (jComboBox10.getSelectedIndex() == 0)) {
            CompleteButton9.setEnabled(false);
        }
        if (jComboBox9.getSelectedIndex() != 0) {
            selected9 = true;
        } else {
            selected9 = false;
        }
    }//GEN-LAST:event_jComboBox9ItemStateChanged

    private void jComboBox10ItemStateChanged(java.awt.event.ItemEvent evt) {//GEN-FIRST:event_jComboBox10ItemStateChanged
        // TODO add your handling code here:
        if ((complete8 == true) && (complete9 == false) && ((jComboBox10.getSelectedIndex() == 0) || (jComboBox9.getSelectedIndex() == 0))) {
            CompleteButton9.setEnabled(false);
        }
        if ((complete8 == true) && (complete9 == false) && (jComboBox10.getSelectedIndex() != 0) && (jComboBox9.getSelectedIndex() != 0)) {
            CompleteButton9.setEnabled(true);
        }
        if (jComboBox10.getSelectedIndex() != 0) {
            selected10 = true;
        } else {
            selected10 = false;
        }


    }//GEN-LAST:event_jComboBox10ItemStateChanged

    public void complete2() {
        complete2 = true;
        completeButtonsPressed.add("CompleteButton2");
        if ((jComboBox3.getSelectedIndex() != 0) && (jComboBox4.getSelectedIndex() != 0)) {
            CompleteButton2.setEnabled(false);
            CompleteButton3.setEnabled(true);
        } else {
            CompleteButton2.setEnabled(false);
        }
        jComboBox2.setEnabled(false);
        upButton2.setEnabled(false);
        downButton2.setEnabled(false);
        ViewButton2.setEnabled(true);
        upButton2.setEnabled(false);
        currentExcelIndex = 2;
        int SelectedIndex = 0;
        String FileItem = (String) jComboBox2.getSelectedItem();
        int FP = 0;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(jComboBox2.getSelectedItem().toString())) {
                FP = i;
            }
        }
        System.out.println(FilePath.get(FP));
        String NameOfFile = FilePath.get(FP) + time.toString() + ".xlsx";
        viewFiles.add(NameOfFile);
        for (int i = 0; i < FileName.size(); i++) {
            if (FileItem.equalsIgnoreCase(FileName.get(i))) {
                SelectedIndex = i;
            }
        }
        addKitReturn(SelectedIndex);
    }

    private void CompleteButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_CompleteButton2ActionPerformed
        // TODO add your handling code here:
        complete2();

    }//GEN-LAST:event_CompleteButton2ActionPerformed

    private void jComboBox3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jComboBox3ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jComboBox3ActionPerformed

    private void CompleteButton1KeyPressed(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_CompleteButton1KeyPressed
        // TODO add your handling code here:
        pressedComplete1 = true;
    }//GEN-LAST:event_CompleteButton1KeyPressed

    private void CompleteButton1StateChanged(javax.swing.event.ChangeEvent evt) {//GEN-FIRST:event_CompleteButton1StateChanged
        // TODO add your handling code here:

    }//GEN-LAST:event_CompleteButton1StateChanged

    private void ViewButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ViewButton1ActionPerformed
        int FP = 0;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(jComboBox1.getSelectedItem().toString())) {
                FP = i;
            }
        }
        System.out.println(FilePath.get(FP));
        String NameOfFile = FilePath.get(FP) + time.toString() + ".xlsx";
        try {
            // TODO add your handling code here:
            Runtime.getRuntime().exec("excel " + NameOfFile);
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_ViewButton1ActionPerformed

    public void complete3() {
        complete3 = true;
        completeButtonsPressed.add("CompleteButton3");
        if ((jComboBox4.getSelectedIndex() != 0) && (jComboBox5.getSelectedIndex() != 0)) {
            CompleteButton3.setEnabled(false);
            CompleteButton4.setEnabled(true);
        } else {
            CompleteButton3.setEnabled(false);
        }
        jComboBox3.setBackground(Color.red);
        jComboBox3.setEnabled(false);
        upButton3.setEnabled(false);
        downButton3.setEnabled(false);
        ViewButton3.setEnabled(true);
        currentExcelIndex = 3;
        int SelectedIndex = 0;
        String FileItem = (String) jComboBox3.getSelectedItem();
        int FP = 0;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(jComboBox3.getSelectedItem().toString())) {
                FP = i;
            }
        }
        System.out.println(FilePath.get(FP));
        String NameOfFile = FilePath.get(FP) + time.toString() + ".xlsx";
        viewFiles.add(NameOfFile);
        for (int i = 0; i < FileName.size(); i++) {
            if (FileItem.equalsIgnoreCase(FileName.get(i))) {
                SelectedIndex = i;
            }
        }
        addKitReturn(SelectedIndex);
    }

    private void CompleteButton3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_CompleteButton3ActionPerformed
        // TODO add your handling code here:
        complete3();

    }//GEN-LAST:event_CompleteButton3ActionPerformed

    private void complete5() {
        complete5 = true;
        completeButtonsPressed.add("CompleteButton5");
        if ((jComboBox6.getSelectedIndex() != 0) && (jComboBox7.getSelectedIndex() != 0)) {
            CompleteButton5.setEnabled(false);
            CompleteButton6.setEnabled(true);
        } else {
            CompleteButton5.setEnabled(false);
        }
        jComboBox5.setBackground(Color.red);
        jComboBox5.setEnabled(false);
        upButton5.setEnabled(false);
        downButton5.setEnabled(false);
        ViewButton5.setEnabled(true);
        currentExcelIndex = 5;
        SelectedFilesPath.clear();
        currentMPN.clear();
        nextMPN.clear();
        indexofcommon.clear();
        System.out.println("selected file path after emptying it" + SelectedFilesPath);
        //the index of item in the filepath of the selected item in the combobox
        currentExcelIndex = 5;
        int SelectedIndex = 0;
        String FileItem = (String) jComboBox5.getSelectedItem();
        int FP = 0;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(jComboBox5.getSelectedItem().toString())) {
                FP = i;
            }
        }
        System.out.println(FilePath.get(FP));
        String NameOfFile = FilePath.get(FP) + time.toString() + ".xlsx";
        viewFiles.add(NameOfFile);
        for (int i = 0; i < FileName.size(); i++) {
            if (FileItem.equalsIgnoreCase(FileName.get(i))) {
                SelectedIndex = i;
            }
        }
        addKitReturn(SelectedIndex);
    }

    private void CompleteButton5ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_CompleteButton5ActionPerformed
        // TODO add your handling code here:
        complete5();
    }//GEN-LAST:event_CompleteButton5ActionPerformed

    private void complete6() {
        complete6 = true;
        completeButtonsPressed.add("CompleteButton6");
        if ((jComboBox7.getSelectedIndex() != 0) && (jComboBox8.getSelectedIndex() != 0)) {
            CompleteButton6.setEnabled(false);
            CompleteButton7.setEnabled(true);
        } else {
            CompleteButton6.setEnabled(false);
        }
        jComboBox6.setBackground(Color.red);
        jComboBox6.setEnabled(false);
        upButton6.setEnabled(false);
        downButton6.setEnabled(false);
        ViewButton6.setEnabled(true);
        currentExcelIndex = 6;
        int SelectedIndex = 0;
        String FileItem = (String) jComboBox6.getSelectedItem();
        int FP = 0;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(jComboBox6.getSelectedItem().toString())) {
                FP = i;
            }
        }
        System.out.println(FilePath.get(FP));
        String NameOfFile = FilePath.get(FP) + time.toString() + ".xlsx";
        viewFiles.add(NameOfFile);
        for (int i = 0; i < FileName.size(); i++) {
            if (FileItem.equalsIgnoreCase(FileName.get(i))) {
                SelectedIndex = i;
            }
        }
        addKitReturn(SelectedIndex);
    }

    private void CompleteButton6ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_CompleteButton6ActionPerformed
        // TODO add your handling code here:
        complete6();

    }//GEN-LAST:event_CompleteButton6ActionPerformed

    private void complete7() {
        complete7 = true;
        completeButtonsPressed.add("CompleteButton7");
        if ((jComboBox8.getSelectedIndex() != 0) && (jComboBox9.getSelectedIndex() != 0)) {
            CompleteButton7.setEnabled(false);
            CompleteButton8.setEnabled(true);
        } else {
            CompleteButton7.setEnabled(false);
        }
        jComboBox7.setBackground(Color.red);
        jComboBox7.setEnabled(false);
        upButton7.setEnabled(false);
        downButton7.setEnabled(false);
        ViewButton7.setEnabled(true);
        currentExcelIndex = 7;
        int SelectedIndex = 0;
        String FileItem = (String) jComboBox7.getSelectedItem();
        int FP = 0;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(jComboBox7.getSelectedItem().toString())) {
                FP = i;
            }
        }
        System.out.println(FilePath.get(FP));
        String NameOfFile = FilePath.get(FP) + time.toString() + ".xlsx";
        viewFiles.add(NameOfFile);
        for (int i = 0; i < FileName.size(); i++) {
            if (FileItem.equalsIgnoreCase(FileName.get(i))) {
                SelectedIndex = i;
            }
        }
        addKitReturn(SelectedIndex);

    }

    private void CompleteButton7ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_CompleteButton7ActionPerformed
        // TODO add your handling code here:
        complete7();
    }//GEN-LAST:event_CompleteButton7ActionPerformed

    private void complete8() {
        complete8 = true;
        completeButtonsPressed.add("CompleteButton8");
        if ((jComboBox9.getSelectedIndex() != 0) && (jComboBox10.getSelectedIndex() != 0)) {
            CompleteButton8.setEnabled(false);
            CompleteButton9.setEnabled(true);
        } else {
            CompleteButton8.setEnabled(false);
        }
        jComboBox8.setBackground(Color.red);
        jComboBox8.setEnabled(false);
        upButton8.setEnabled(false);
        downButton8.setEnabled(false);
        ViewButton8.setEnabled(true);
        currentExcelIndex = 8;
        int SelectedIndex = 0;
        String FileItem = (String) jComboBox8.getSelectedItem();
        int FP = 0;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(jComboBox8.getSelectedItem().toString())) {
                FP = i;
            }
        }
        System.out.println(FilePath.get(FP));
        String NameOfFile = FilePath.get(FP) + time.toString() + ".xlsx";
        viewFiles.add(NameOfFile);
        for (int i = 0; i < FileName.size(); i++) {
            if (FileItem.equalsIgnoreCase(FileName.get(i))) {
                SelectedIndex = i;
            }
        }
        addKitReturn(SelectedIndex);

    }

    private void CompleteButton8ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_CompleteButton8ActionPerformed
        // TODO add your handling code here:
        complete8();
    }//GEN-LAST:event_CompleteButton8ActionPerformed

    private void complete9() {
        complete9 = true;
        completeButtonsPressed.add("CompleteButton9");
        if (jComboBox10.getSelectedIndex() != 0) {
            CompleteButton9.setEnabled(false);
            ViewButton10.setEnabled(true);
        }
        jComboBox9.setBackground(Color.red);
        jComboBox9.setEnabled(false);
        upButton9.setEnabled(false);
        downButton9.setEnabled(false);
        ViewButton9.setEnabled(true);
        currentExcelIndex = 9;
        int SelectedIndex = 0;
        String FileItem = (String) jComboBox9.getSelectedItem();
        int FP = 0;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(jComboBox9.getSelectedItem().toString())) {
                FP = i;
            }
        }
        System.out.println(FilePath.get(FP));
        String NameOfFile = FilePath.get(FP) + time.toString() + ".xlsx";
        viewFiles.add(NameOfFile);
        for (int i = 0; i < FileName.size(); i++) {
            if (FileItem.equalsIgnoreCase(FileName.get(i))) {
                SelectedIndex = i;
            }
        }
        addKitReturn(SelectedIndex);
    }

    private void CompleteButton9ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_CompleteButton9ActionPerformed
        // TODO add your handling code here:
        complete9();

    }//GEN-LAST:event_CompleteButton9ActionPerformed

    private void downButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_downButton1ActionPerformed
        // TODO add your handling code here:
        int i = jComboBox1.getSelectedIndex();
        int j = jComboBox2.getSelectedIndex();
        jComboBox2.setSelectedIndex(i);
        jComboBox1.setSelectedIndex(j);

    }//GEN-LAST:event_downButton1ActionPerformed

    private void upButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_upButton1ActionPerformed
        // TODO add your handling code here:
        int i = jComboBox1.getSelectedIndex();
        int j = jComboBox2.getSelectedIndex();
        jComboBox2.setSelectedIndex(i);
        jComboBox1.setSelectedIndex(j);

    }//GEN-LAST:event_upButton1ActionPerformed

    private void upButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_upButton2ActionPerformed
        // TODO add your handling code here:
        int i = jComboBox2.getSelectedIndex();
        int j = jComboBox3.getSelectedIndex();
        jComboBox3.setSelectedIndex(i);
        jComboBox2.setSelectedIndex(j);
    }//GEN-LAST:event_upButton2ActionPerformed

    private void downButton3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_downButton3ActionPerformed
        // TODO add your handling code here:
        int i = jComboBox3.getSelectedIndex();
        int j = jComboBox4.getSelectedIndex();
        jComboBox4.setSelectedIndex(i);
        jComboBox3.setSelectedIndex(j);
    }//GEN-LAST:event_downButton3ActionPerformed

    private void upButton3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_upButton3ActionPerformed
        // TODO add your handling code here:
        int i = jComboBox3.getSelectedIndex();
        int j = jComboBox4.getSelectedIndex();
        jComboBox4.setSelectedIndex(i);
        jComboBox3.setSelectedIndex(j);
    }//GEN-LAST:event_upButton3ActionPerformed

    private void downButton4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_downButton4ActionPerformed
        // TODO add your handling code here:
        int i = jComboBox4.getSelectedIndex();
        int j = jComboBox5.getSelectedIndex();
        jComboBox5.setSelectedIndex(i);
        jComboBox4.setSelectedIndex(j);
    }//GEN-LAST:event_downButton4ActionPerformed

    private void upButton4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_upButton4ActionPerformed
        // TODO add your handling code here:
        int i = jComboBox4.getSelectedIndex();
        int j = jComboBox5.getSelectedIndex();
        jComboBox5.setSelectedIndex(i);
        jComboBox4.setSelectedIndex(j);
    }//GEN-LAST:event_upButton4ActionPerformed

    private void downButton5ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_downButton5ActionPerformed
        // TODO add your handling code here:
        int i = jComboBox5.getSelectedIndex();
        int j = jComboBox6.getSelectedIndex();
        jComboBox6.setSelectedIndex(i);
        jComboBox5.setSelectedIndex(j);
    }//GEN-LAST:event_downButton5ActionPerformed

    private void upButton5ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_upButton5ActionPerformed
        // TODO add your handling code here:
        int i = jComboBox5.getSelectedIndex();
        int j = jComboBox6.getSelectedIndex();
        jComboBox6.setSelectedIndex(i);
        jComboBox5.setSelectedIndex(j);
    }//GEN-LAST:event_upButton5ActionPerformed

    private void downButton6ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_downButton6ActionPerformed
        // TODO add your handling code here:
        int i = jComboBox6.getSelectedIndex();
        int j = jComboBox7.getSelectedIndex();
        jComboBox7.setSelectedIndex(i);
        jComboBox6.setSelectedIndex(j);
    }//GEN-LAST:event_downButton6ActionPerformed

    private void upButton6ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_upButton6ActionPerformed
        // TODO add your handling code here:
        int i = jComboBox6.getSelectedIndex();
        int j = jComboBox7.getSelectedIndex();
        jComboBox7.setSelectedIndex(i);
        jComboBox6.setSelectedIndex(j);
    }//GEN-LAST:event_upButton6ActionPerformed

    private void upButton7ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_upButton7ActionPerformed
        // TODO add your handling code here:
        int i = jComboBox7.getSelectedIndex();
        int j = jComboBox8.getSelectedIndex();
        jComboBox8.setSelectedIndex(i);
        jComboBox7.setSelectedIndex(j);
    }//GEN-LAST:event_upButton7ActionPerformed

    private void downButton8ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_downButton8ActionPerformed
        // TODO add your handling code here:
        int i = jComboBox9.getSelectedIndex();
        int j = jComboBox8.getSelectedIndex();
        jComboBox9.setSelectedIndex(i);
        jComboBox8.setSelectedIndex(j);
    }//GEN-LAST:event_downButton8ActionPerformed

    private void upButton8ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_upButton8ActionPerformed
        // TODO add your handling code here:
        int i = jComboBox8.getSelectedIndex();
        int j = jComboBox9.getSelectedIndex();
        jComboBox9.setSelectedIndex(i);
        jComboBox8.setSelectedIndex(j);
    }//GEN-LAST:event_upButton8ActionPerformed

    private void downButton9ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_downButton9ActionPerformed
        // TODO add your handling code here:
        int i = jComboBox9.getSelectedIndex();
        int j = jComboBox10.getSelectedIndex();
        jComboBox10.setSelectedIndex(i);
        jComboBox9.setSelectedIndex(j);
    }//GEN-LAST:event_downButton9ActionPerformed

    private void upButton9ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_upButton9ActionPerformed
        // TODO add your handling code here:
        int i = jComboBox9.getSelectedIndex();
        int j = jComboBox10.getSelectedIndex();
        jComboBox10.setSelectedIndex(i);
        jComboBox9.setSelectedIndex(j);
    }//GEN-LAST:event_upButton9ActionPerformed

    private void ViewButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ViewButton2ActionPerformed
        // TODO add your handling code here:
        int FP = 0;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(jComboBox2.getSelectedItem().toString())) {
                FP = i;
            }
        }
        String NameOfFile = FilePath.get(FP) + time.toString() + ".xlsx";
        try {
            // TODO add your handling code here:
            Runtime.getRuntime().exec("excel " + NameOfFile);
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_ViewButton2ActionPerformed

    private void ViewButton3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ViewButton3ActionPerformed
        // TODO add your handling code here:
        int FP = 0;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(jComboBox3.getSelectedItem().toString())) {
                FP = i;
            }
        }
        String NameOfFile = FilePath.get(FP) + time.toString() + ".xlsx";
        try {
            // TODO add your handling code here:
            Runtime.getRuntime().exec("excel " + NameOfFile);
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_ViewButton3ActionPerformed

    private void ViewButton6ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ViewButton6ActionPerformed
        // TODO add your handling code here:
        int FP = 0;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(jComboBox6.getSelectedItem().toString())) {
                FP = i;
            }
        }
        String NameOfFile = FilePath.get(FP) + time.toString() + ".xlsx";
        try {
            // TODO add your handling code here:
            Runtime.getRuntime().exec("excel " + NameOfFile);
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_ViewButton6ActionPerformed

    private void ViewButton7ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ViewButton7ActionPerformed
        // TODO add your handling code here:
        int FP = 0;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(jComboBox7.getSelectedItem().toString())) {
                FP = i;
            }
        }
        String NameOfFile = FilePath.get(FP) + time.toString() + ".xlsx";
        try {
            // TODO add your handling code here:
            Runtime.getRuntime().exec("excel " + NameOfFile);
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_ViewButton7ActionPerformed

    private void ViewButton8ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ViewButton8ActionPerformed
        // TODO add your handling code here:
        int FP = 0;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(jComboBox8.getSelectedItem().toString())) {
                FP = i;
            }
        }
        String NameOfFile = FilePath.get(FP) + time.toString() + ".xlsx";
        try {
            // TODO add your handling code here:
            Runtime.getRuntime().exec("excel " + NameOfFile);
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_ViewButton8ActionPerformed

    private void ViewButton9ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ViewButton9ActionPerformed
        // TODO add your handling code here:
        int FP = 0;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(jComboBox9.getSelectedItem().toString())) {
                FP = i;
            }
        }
        String NameOfFile = FilePath.get(FP) + time.toString() + ".xlsx";
        try {
            // TODO add your handling code here:
            Runtime.getRuntime().exec("excel " + NameOfFile);
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_ViewButton9ActionPerformed

    private void ViewButton10ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ViewButton10ActionPerformed
        // TODO add your handling code here:
        int FP = 0;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(jComboBox10.getSelectedItem().toString())) {
                FP = i;
            }
        }
        String NameOfFile = FilePath.get(FP) + time.toString() + ".xlsx";
        try {
            // TODO add your handling code here:
            Runtime.getRuntime().exec("excel " + NameOfFile);
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_ViewButton10ActionPerformed

    private void CompleteButton2KeyPressed(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_CompleteButton2KeyPressed
        // TODO add your handling code here:
        complete2 = true;
    }//GEN-LAST:event_CompleteButton2KeyPressed

    private void SaveButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_SaveButtonActionPerformed
        try {
            // TODO add your handling code here:
            saveFile();
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_SaveButtonActionPerformed

    public boolean getFinish() {
        return finish;
    }
    //the files that have been selected in the combo boxes
    private ArrayList<String> SelectedFilesPath = new ArrayList<>();
    //Arraylist that included the manufacter part number column of the current project
    private ArrayList<String> currentMPN = new ArrayList<>();
    private ArrayList<String> nextMPN = new ArrayList<>();
    private ArrayList<String> commonMPN = new ArrayList<>();
    private ArrayList<String> OpenSelectedComboBoxes = new ArrayList<>();
    //current excel file/project that is happening 
    int currentExcelIndex = -1;
    //whether the complete button has been pressed
    private boolean pressedComplete1 = false;
    //the indexes of where the common manufacturer part numbers of the two projects exist in the current project
    private HashSet<Integer> indexofcommon = new HashSet<Integer>();
    //the files that can be viewed
    private ArrayList<String> viewFiles = new ArrayList<>();
    //complete buttons pressed
    private ArrayList<String> completeButtonsPressed = new ArrayList<>();
    //whether the complete button should be enabled
    private boolean complete2 = false;
    private boolean complete3 = false;
    private boolean complete4 = false;
    private boolean complete5 = false;
    private boolean complete6 = false;
    private boolean complete7 = false;
    private boolean complete8 = false;
    private boolean complete9 = false;
    private boolean complete10 = false;
    //whether an item is selected in combo box2
    private boolean selected2 = false;
    private boolean selected3 = false;
    private boolean selected4 = false;
    private boolean selected5 = false;
    private boolean selected6 = false;
    private boolean selected7 = false;
    private boolean selected8 = false;
    private boolean selected9 = false;
    private boolean selected10 = false;
    private LocalDate time = java.time.LocalDate.now();
    // It is a boolean that enables the finish button
    private boolean finish = false;
    private String currentFilePath;
    // an arraylist with the path of the excel files
    private ArrayList<String> FilePath = new ArrayList<>();
    //It is an index for 
    //private int index = 0;
    // an arraylist with the names of the excel files
    private ArrayList<String> FileName = new ArrayList<>();
    private boolean start = false;
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton AddExcelButton;
    private javax.swing.JButton CompleteButton1;
    private javax.swing.JButton CompleteButton2;
    private javax.swing.JButton CompleteButton3;
    private javax.swing.JButton CompleteButton4;
    private javax.swing.JButton CompleteButton5;
    private javax.swing.JButton CompleteButton6;
    private javax.swing.JButton CompleteButton7;
    private javax.swing.JButton CompleteButton8;
    private javax.swing.JButton CompleteButton9;
    private javax.swing.JButton RemoveExcelButton;
    private javax.swing.JButton SaveButton;
    private javax.swing.JButton ViewButton1;
    private javax.swing.JButton ViewButton10;
    private javax.swing.JButton ViewButton2;
    private javax.swing.JButton ViewButton3;
    private javax.swing.JButton ViewButton4;
    private javax.swing.JButton ViewButton5;
    private javax.swing.JButton ViewButton6;
    private javax.swing.JButton ViewButton7;
    private javax.swing.JButton ViewButton8;
    private javax.swing.JButton ViewButton9;
    private javax.swing.JButton downButton1;
    private javax.swing.JButton downButton2;
    private javax.swing.JButton downButton3;
    private javax.swing.JButton downButton4;
    private javax.swing.JButton downButton5;
    private javax.swing.JButton downButton6;
    private javax.swing.JButton downButton7;
    private javax.swing.JButton downButton8;
    private javax.swing.JButton downButton9;
    private javax.swing.JComboBox<String> jComboBox1;
    private javax.swing.JComboBox<String> jComboBox10;
    private javax.swing.JComboBox<String> jComboBox2;
    private javax.swing.JComboBox<String> jComboBox3;
    private javax.swing.JComboBox<String> jComboBox4;
    private javax.swing.JComboBox<String> jComboBox5;
    private javax.swing.JComboBox<String> jComboBox6;
    private javax.swing.JComboBox<String> jComboBox7;
    private javax.swing.JComboBox<String> jComboBox8;
    private javax.swing.JComboBox<String> jComboBox9;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel10;
    private javax.swing.JLabel jLabel11;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JLabel jLabel8;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JMenu jMenu1;
    private javax.swing.JMenu jMenu2;
    private javax.swing.JMenu jMenu3;
    private javax.swing.JMenu jMenu4;
    private javax.swing.JMenuBar jMenuBar1;
    private javax.swing.JMenuBar jMenuBar2;
    private javax.swing.JButton upButton1;
    private javax.swing.JButton upButton2;
    private javax.swing.JButton upButton3;
    private javax.swing.JButton upButton4;
    private javax.swing.JButton upButton5;
    private javax.swing.JButton upButton6;
    private javax.swing.JButton upButton7;
    private javax.swing.JButton upButton8;
    private javax.swing.JButton upButton9;
    // End of variables declaration//GEN-END:variables
}
