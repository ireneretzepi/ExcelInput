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

    private static final int MAX_FILES = 10;

    private final JComboBox<String>[] comboBoxes = new JComboBox[MAX_FILES];
    private final JButton[] upButtons = new JButton[MAX_FILES];
    private final JButton[] downButtons = new JButton[MAX_FILES];
    private final JButton[] completeButtons = new JButton[MAX_FILES - 1]; // 9 buttons for 10 combos
    private final JButton[] viewButtons = new JButton[MAX_FILES];
    
    private final boolean[] completed = new boolean[MAX_FILES - 1];
    
    private boolean selected2 = false;
    private boolean selected3 = false;
    private boolean selected4 = false;
    private boolean selected5 = false;
    private boolean selected6 = false;
    private boolean selected7 = false;
    private boolean selected8 = false;
    private boolean selected9 = false;
    private boolean selected10 = false;
    
    private boolean pressedComplete1 = false;
    private int currentExcelIndex = 0;
    private String currentFilePath = "";
    private final LocalTime time = LocalTime.now();
    private boolean start = true;
    
    private final ArrayList<String> FileName = new ArrayList<>();
    private final ArrayList<String> FilePath = new ArrayList<>();
    private final ArrayList<String> SelectedFilesPath = new ArrayList<>();
    private final ArrayList<String> currentMPN = new ArrayList<>();
    private final ArrayList<String> nextMPN = new ArrayList<>();
    private final ArrayList<Integer> indexofcommon = new ArrayList<>();
    private final ArrayList<String> viewFiles = new ArrayList<>();
    private final ArrayList<String> completeButtonsPressed = new ArrayList<>();
    private final JComboBox<String>[] comboBoxes = new JComboBox[MAX_FILES];



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
        if (start) {
            resetComboBoxes();
        }
        start = false;
    
        File selectedFile = promptForExcelFile();
        if (selectedFile == null) return;
    
        String filename = selectedFile.getName();
        if (isDuplicateFile(filename)) {
            JOptionPane.showMessageDialog(this, "The file you chose already exists!");
            return;
        }
    
        readExcel(selectedFile.getAbsolutePath());
    
        FilePath.add(selectedFile.getAbsolutePath());
        FileName.add(filename);
        addFileToComboBoxes(filename);
    
        RemoveExcelButton.setEnabled(true);
    }

    private void resetComboBoxes() {
        JComboBox<?>[] comboBoxes = {
            jComboBox1, jComboBox2, jComboBox3, jComboBox4, jComboBox5,
            jComboBox6, jComboBox7, jComboBox8, jComboBox9, jComboBox10
        };
        for (JComboBox<?> comboBox : comboBoxes) {
            comboBox.setSelectedIndex(0);
        }
    }
    
    private File promptForExcelFile() {
        JFileChooser chooser = new JFileChooser();
        chooser.setCurrentDirectory(new File(System.getProperty("user.home")));
        int returnValue;
    
        while (true) {
            returnValue = chooser.showOpenDialog(this);
            if (returnValue != JFileChooser.APPROVE_OPTION) return null;
    
            File selectedFile = chooser.getSelectedFile();
            String extension = getFileExtension(selectedFile.getName());
    
            if ("xlsx".equalsIgnoreCase(extension)) {
                return selectedFile;
            }
    
            JOptionPane.showMessageDialog(this, "Please select a valid Excel (.xlsx) file.");
        }
    }
    
    private boolean isDuplicateFile(String filename) {
        return FileName.stream().anyMatch(existing -> existing.equalsIgnoreCase(filename));
    }
    
    private String getFileExtension(String filename) {
        int dotIndex = filename.lastIndexOf('.');
        return (dotIndex != -1 && dotIndex < filename.length() - 1) ? filename.substring(dotIndex + 1) : "";
    }
    
    private void addFileToComboBoxes(String filename) {
        JComboBox<String>[] comboBoxes = new JComboBox[]{
            jComboBox1, jComboBox2, jComboBox3, jComboBox4, jComboBox5,
            jComboBox6, jComboBox7, jComboBox8, jComboBox9, jComboBox10
        };
        for (JComboBox<String> comboBox : comboBoxes) {
            comboBox.addItem(filename);
        }
    }



/**
 * Saves the current application state (file names, paths, combo selections, etc.) to a text file.
 */
    public void saveFile() throws IOException {
        JFileChooser chooser = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
        int returnValue = chooser.showSaveDialog(null);
    
        if (returnValue != JFileChooser.APPROVE_OPTION) return;
    
        File file = chooser.getSelectedFile();
    
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(file))) {
            writeSection(writer, "FileName ArrayList:", FileName);
            writeSection(writer, "FilePath ArrayList:", FilePath);
    
            List<String> selectedComboValues = getSelectedComboBoxItems();
            writeSection(writer, "Selected Combo Boxes:", selectedComboValues);
    
            writer.write("Number of up/down buttons that should be disabled:" + completeButtonsPressed.size());
            writer.newLine();
    
            writeSection(writer, "View files created:", viewFiles);
            writeSection(writer, "Complete Buttons Pressed:", completeButtonsPressed);
        }
    }

    private void writeSection(BufferedWriter writer, String title, List<String> values) throws IOException {
        writer.write(title);
        writer.newLine();
        for (String value : values) {
            writer.write(value);
            writer.newLine();
        }
    }
    
    private List<String> getSelectedComboBoxItems() {
        JComboBox<?>[] comboBoxes = {
            jComboBox1, jComboBox2, jComboBox3, jComboBox4, jComboBox5,
            jComboBox6, jComboBox7, jComboBox8, jComboBox9, jComboBox10
        };
    
        List<String> selectedItems = new ArrayList<>();
        for (JComboBox<?> comboBox : comboBoxes) {
            int index = comboBox.getSelectedIndex();
            if (index > 0 && comboBox.getSelectedItem() != null) {
                selectedItems.add(comboBox.getSelectedItem().toString());
            }
        }
        return selectedItems;
    }


/**
 * Restores the previous GUI state using the provided saved data.
 */
    public void readFile(
            List<String> fileNameData,
            List<String> filePathData,
            List<String> selectedComboBoxItems,
            String disableUpDown,
            List<String> viewFilesData,
            List<String> completedButtonsData
    ) {
        populateFileData(fileNameData, filePathData);
        OpenSelectedComboBoxes.addAll(selectedComboBoxItems);
        viewFiles.addAll(viewFilesData);
        completeButtonsPressed = completedButtonsData;
    
        List<Integer> selectedIndices = mapComboSelections(OpenSelectedComboBoxes);
        applyComboSelections(selectedIndices);
        restoreCompletedButtons(completedButtonsData, viewFilesData);
    
        System.out.println("OpenSelectedComboBoxes: " + OpenSelectedComboBoxes);
        System.out.println("indexOfSelectedItem: " + selectedIndices);
    }

    private void populateFileData(List<String> fileNames, List<String> filePaths) {
        for (int i = 0; i < fileNames.size(); i++) {
            String name = fileNames.get(i);
            FileName.add(name);
            FilePath.add(filePaths.get(i));
            addFileToComboBoxes(name);
        }
    }
    
    private void addFileToComboBoxes(String filename) {
        JComboBox<String>[] comboBoxes = getComboBoxes();
        for (JComboBox<String> comboBox : comboBoxes) {
            comboBox.addItem(filename);
        }
    }
    
    private List<Integer> mapComboSelections(List<String> selectedNames) {
        List<Integer> indices = new ArrayList<>();
        for (String selected : selectedNames) {
            for (int i = 0; i < FileName.size(); i++) {
                if (FileName.get(i).equalsIgnoreCase(selected)) {
                    indices.add(i + 1); // 1-based index
                    break;
                }
            }
        }
        return indices;
    }
    
    private void applyComboSelections(List<Integer> indices) {
        JComboBox<String>[] comboBoxes = getComboBoxes();
        for (int i = 0; i < indices.size() && i < comboBoxes.length; i++) {
            comboBoxes[i].setSelectedIndex(indices.get(i));
        }
    }
    
    private void restoreCompletedButtons(List<String> completedButtons, List<String> viewFilesList) {
        for (int i = 0; i < completedButtons.size(); i++) {
            String buttonId = completedButtons.get(i);
            Runnable completeAction = getCompleteAction(buttonId);
            JButton viewButton = getViewButton(buttonId);
    
            if (completeAction != null && viewButton != null) {
                completeAction.run();
                viewButton.setEnabled(true);
                launchExcel(viewFilesList.get(i));
            }
        }
    }
    
    private Runnable getCompleteAction(String buttonId) {
        return switch (buttonId) {
            case "CompleteButton1" -> this::Complete1;
            case "CompleteButton2" -> this::complete2;
            case "CompleteButton3" -> this::complete3;
            case "CompleteButton4" -> this::complete4;
            case "CompleteButton5" -> this::complete5;
            case "CompleteButton6" -> this::complete6;
            case "CompleteButton7" -> this::complete7;
            case "CompleteButton8" -> this::complete8;
            case "CompleteButton9" -> this::complete9;
            default -> null;
        };
    }
    
    private JButton getViewButton(String buttonId) {
        return switch (buttonId) {
            case "CompleteButton1" -> ViewButton1;
            case "CompleteButton2" -> ViewButton2;
            case "CompleteButton3" -> ViewButton3;
            case "CompleteButton4" -> ViewButton4;
            case "CompleteButton5" -> ViewButton5;
            case "CompleteButton6" -> ViewButton6;
            case "CompleteButton7" -> ViewButton7;
            case "CompleteButton8" -> ViewButton8;
            case "CompleteButton9" -> ViewButton9;
            default -> null;
        };
    }
    
    private void launchExcel(String path) {
        try {
            Runtime.getRuntime().exec("excel " + path);
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    @SuppressWarnings("unchecked")
    private JComboBox<String>[] getComboBoxes() {
        return new JComboBox[]{
            jComboBox1, jComboBox2, jComboBox3, jComboBox4, jComboBox5,
            jComboBox6, jComboBox7, jComboBox8, jComboBox9, jComboBox10
        };
    }


    public ArrayList getFilePath() {
        return FilePath;
    }

    public void readExcel(String testFileName) {

    }

   private void initComponents() {
        setOpaque(false);
        setPreferredSize(new Dimension(1000, 800));

        JLabel headerLabel = new JLabel("Input the Excel files in the order the projects will be completed!");
        JButton addExcelButton = new JButton("Add Excel File");
        JButton removeExcelButton = new JButton("Remove Excel File");
        JButton saveButton = new JButton("Save");

        JPanel topPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        topPanel.add(headerLabel);
        topPanel.add(addExcelButton);

        JPanel mainPanel = new JPanel();
        mainPanel.setLayout(new BoxLayout(mainPanel, BoxLayout.Y_AXIS));

        for (int i = 0; i < MAX_FILES; i++) {
            comboBoxes[i] = new JComboBox<>(new String[]{"Select Excel File"});
            comboBoxes[i].setOpaque(false);

            upButtons[i] = new JButton("↑");
            downButtons[i] = new JButton("↓");
            completeButtons[i] = new JButton("Complete");
            viewButtons[i] = new JButton("View👁");

            JPanel row = new JPanel();
            row.setLayout(new FlowLayout(FlowLayout.LEFT));
            row.add(new JLabel((i + 1) + "."));
            row.add(comboBoxes[i]);
            row.add(upButtons[i]);
            row.add(downButtons[i]);
            row.add(completeButtons[i]);
            row.add(viewButtons[i]);

            mainPanel.add(row);
        }

        JPanel bottomPanel = new JPanel();
        bottomPanel.add(removeExcelButton);
        bottomPanel.add(saveButton);

        setLayout(new BorderLayout());
        add(topPanel, BorderLayout.NORTH);
        add(mainPanel, BorderLayout.CENTER);
        add(bottomPanel, BorderLayout.SOUTH);
    }
    

    private void AddExcelButtonActionPerformed(java.awt.event.ActionEvent evt) {
        addExcel(); // This should contain your file chooser + update logic
    }

    private void RemoveExcelButtonActionPerformed(java.awt.event.ActionEvent evt) {
        if (FileName.isEmpty()) return;
    
        String[] options = FileName.toArray(new String[0]);
        String input = (String) JOptionPane.showInputDialog(
            null,
            "Choose the name of the Excel file you want to remove:",
            "Choose remove file name",
            JOptionPane.QUESTION_MESSAGE,
            null,
            options,
            options[0]
        );
    
        if (input == null) return;
    
        int index = -1;
        for (int i = 0; i < FileName.size(); i++) {
            if (FileName.get(i).equalsIgnoreCase(input)) {
                index = i;
                break;
            }
        }
    
        if (index != -1) {
            FileName.remove(index);
            FilePath.remove(index);
            for (JComboBox<String> comboBox : comboBoxes) {
                if (comboBox.getItemCount() > index + 1) {
                    comboBox.removeItemAt(index + 1); // +1 because "Select Excel File" is at index 0
                }
            }
        }
    
        RemoveExcelButton.setEnabled(!FileName.isEmpty());
    }


    private void initComboBoxListeners() {
        for (int i = 0; i < comboBoxes.length; i++) {
            final int index = i;
            comboBoxes[i].addActionListener(e -> onComboBoxSelected(index, e));
        }
    }
    
    private void onComboBoxSelected(int index, ActionEvent e) {
        // You can add logging, validation, or tracking here
        System.out.println("ComboBox " + (index + 1) + " selected: " + comboBoxes[index].getSelectedItem());
    }

    private void swapComboBoxSelection(JComboBox<String> boxA, JComboBox<String> boxB) {
        int indexA = boxA.getSelectedIndex();
        int indexB = boxB.getSelectedIndex();
        boxA.setSelectedIndex(indexB);
        boxB.setSelectedIndex(indexA);
    }


    private void downButton2ActionPerformed(ActionEvent evt) {
        swapComboBoxSelection(jComboBox2, jComboBox3);
    }
    
    private void downButton7ActionPerformed(ActionEvent evt) {
        swapComboBoxSelection(jComboBox7, jComboBox8);
    }


    private void complete(int index) {
        completeButtonsPressed.add("CompleteButton" + (index + 1));
        currentExcelIndex = index + 1;
    
        if (index < comboBoxes.length - 2 &&
            comboBoxes[index + 1].getSelectedIndex() != 0 &&
            comboBoxes[index + 2].getSelectedIndex() != 0) {
            completeButtons[index].setEnabled(false);
            completeButtons[index + 1].setEnabled(true);
        } else {
            completeButtons[index].setEnabled(false);
        }
    
        comboBoxes[index].setBackground(Color.RED);
        comboBoxes[index].setEnabled(false);
        upButtons[index].setEnabled(false);
        downButtons[index].setEnabled(false);
        viewButtons[index].setEnabled(true);
        RemoveExcelButton.setEnabled(true);
        SaveButton.setEnabled(true);
    
        String selectedItem = (String) comboBoxes[index].getSelectedItem();
        int fileIndex = FileName.indexOf(selectedItem);
    
        if (fileIndex != -1) {
            String nameOfFile = FilePath.get(fileIndex) + time.toString() + ".xlsx";
            viewFiles.add(nameOfFile);
            addKitReturn(fileIndex);
            System.out.println(FilePath.get(fileIndex));
        }
    }


    private void disableComboBox(JComboBox<String> comboBox) {
        comboBox.setBackground(Color.RED);
        comboBox.setEnabled(false);
    }


   private void CompleteButton4ActionPerformed(java.awt.event.ActionEvent evt) {
        complete(3);
    }


    private void ViewButton4ActionPerformed(java.awt.event.ActionEvent evt) {
        viewFileFromComboBox(jComboBox4);
    }
    
    private void ViewButton5ActionPerformed(java.awt.event.ActionEvent evt) {
        viewFileFromComboBox(jComboBox5);
    }

    private void viewFileFromComboBox(JComboBox<String> comboBox) {
        String selected = (String) comboBox.getSelectedItem();
        if (selected == null) return;
    
        int index = FileName.indexOf(selected);
        if (index == -1) return;
    
        String fullPath = FilePath.get(index) + time.toString() + ".xlsx";
        System.out.println(fullPath);
        try {
            Runtime.getRuntime().exec("excel \"" + fullPath + "\"");
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private void CompleteButton1ActionPerformed(ActionEvent evt) {
        complete(0);
    }


    public void addKitReturn(int selectedIndex) {
        // Clear previous state
        SelectedFilesPath.clear();
        currentMPN.clear();
        nextMPN.clear();
        indexofcommon.clear();
    
        System.out.println("Selected file path after clearing: " + SelectedFilesPath);
        createPathArraylist();
    
        currentFilePath = FilePath.get(selectedIndex);
        System.out.println("Current file path: " + currentFilePath);
    
        try (FileInputStream excelFile = new FileInputStream(new File(currentFilePath))) {
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();
    
            // Set style for "Re-Kit" text
            XSSFFont font = (XSSFFont) workbook.createFont();
            font.setBold(true);
            font.setUnderline(FontUnderline.SINGLE);
            XSSFCellStyle style = (XSSFCellStyle) workbook.createCellStyle();
            style.setFont(font);
    
            boolean foundHeader = false;
    
            for (Row row : sheet) {
                Cell cell = row.getCell(4);
                if (cell == null) continue;
    
                String cellValue = formatter.formatCellValue(cell);
                if (foundHeader && !cellValue.isEmpty()) {
                    currentMPN.add(cellValue);
                }
                if ("Manufacturer Part Number".equalsIgnoreCase(cellValue)) {
                    foundHeader = true;
                }
            }
    
            for (Row row : sheet) {
                row.createCell(6);  // Create column where "Re-Kit" may go
    
                for (Cell cell : row) {
                    String value = formatter.formatCellValue(cell);
                    if ("Populate".equalsIgnoreCase(value)) {
                        if (row.getLastCellNum() > 6) {
                            row.getCell(6).setCellValue("Re-Kit");
                            row.getCell(6).setCellStyle(style);
                        }
                        break;
                    }
                }
            }
    
            String newFilename = currentFilePath + java.time.LocalDate.now() + ".xlsx";
            try (FileOutputStream outputFile = new FileOutputStream(new File(newFilename))) {
                workbook.write(outputFile);
            }
    
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }
    
        for (int i = 0; i < SelectedFilesPath.size(); i++) {
            compare(i);
        }
    }


    public void compare(int index) {
        nextMPN.clear();
        String nextFileName = "";
    
        String selectedPath = SelectedFilesPath.get(index);
    
        // Find the matching filename
        for (int i = 0; i < FilePath.size(); i++) {
            if (FilePath.get(i).equals(selectedPath)) {
                nextFileName = FileName.get(i);
                break;
            }
        }
    
        System.out.println("Next file: " + nextFileName);
        System.out.println("Selected path at index: " + selectedPath);
    
        // Extract MPNs from next file
        try (FileInputStream fis = new FileInputStream(selectedPath)) {
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();
            boolean foundHeader = false;
    
            for (Row row : sheet) {
                Cell cell = row.getCell(4);
                if (cell == null) continue;
                String value = formatter.formatCellValue(cell);
    
                if (foundHeader && !value.isEmpty()) {
                    nextMPN.add(value);
                }
                if ("Manufacturer Part Number".equalsIgnoreCase(value)) {
                    foundHeader = true;
                }
            }
    
            // Save (even if unchanged)
            try (FileOutputStream fos = new FileOutputStream(selectedPath)) {
                workbook.write(fos);
            }
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }
    
        // Compare MPN lists
        List<String> current = new ArrayList<>(currentMPN);
        List<String> common = new ArrayList<>(currentMPN);
        current.removeAll(nextMPN);
        common.removeAll(current);
    
        System.out.println("Common MPNs: " + common);
        System.out.println("Current file path: " + currentFilePath);
    
        // Add file name to cells with common MPNs
        try (FileInputStream fis = new FileInputStream(currentFilePath + java.time.LocalDate.now() + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();
            boolean foundHeader = false;
            int commonIndex = 0;
    
            for (Row row : sheet) {
                Cell cell = row.getCell(4);
                if (cell == null) continue;
    
                String value = formatter.formatCellValue(cell);
    
                if (foundHeader && commonIndex < common.size()) {
                    if (value.equalsIgnoreCase(common.get(commonIndex))) {
                        indexofcommon.add(row.getRowNum());
    
                        int newColIndex = 6 + index;
                        Cell cellToUpdate = row.getCell(newColIndex);
                        if (cellToUpdate == null) {
                            cellToUpdate = row.createCell(newColIndex);
                        }
                        cellToUpdate.setCellValue(nextFileName);
                        commonIndex++;
                    }
                }
    
                if ("Manufacturer Part Number".equalsIgnoreCase(value)) {
                    foundHeader = true;
                }
            }
    
            try (FileOutputStream fos = new FileOutputStream(currentFilePath + java.time.LocalDate.now() + ".xlsx")) {
                workbook.write(fos);
            }
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }
    }


    public void createPathArraylist() {
        SelectedFilesPath.clear();
    
        // comboBoxes[1] to comboBoxes[9] correspond to jComboBox2 to jComboBox10
        for (int i = 1; i < comboBoxes.length; i++) {
            if (i - 1 < completed.length && completed[i - 1]) continue;
    
            JComboBox<String> comboBox = comboBoxes[i];
            Object selectedItem = comboBox.getSelectedItem();
            if (selectedItem == null) continue;
    
            String selectedName = selectedItem.toString();
            if (selectedName.equalsIgnoreCase("Select Excel File")) continue;
    
            int fileIndex = FileName.indexOf(selectedName);
            if (fileIndex != -1) {
                SelectedFilesPath.add(FilePath.get(fileIndex));
            }
        }
    }


    /**
     * Updates the enabled state of all Complete buttons based on the
     * completion flags and combo box selections. This ensures buttons
     * are only clickable when both preceding and current selections are valid.
     */
    private void updateComboBoxStates() {
        for (int i = 0; i < completeButtons.length; i++) {
            boolean isPrevComplete = (i == 0) ? completed[0] : completed[i - 1];
            boolean isCurrentComplete = completed[i];
    
            JComboBox comboBoxA = comboBoxes[i];
            JComboBox comboBoxB = comboBoxes[i + 1];
            JButton completeButton = completeButtons[i];
    
            boolean comboBoxASelected = comboBoxA.getSelectedIndex() > 0;
            boolean comboBoxBSelected = comboBoxB.getSelectedIndex() > 0;
    
            boolean shouldEnable = false;
            if (isPrevComplete && comboBoxASelected && comboBoxBSelected) {
                shouldEnable = true;
            }
    
            completeButton.setEnabled(shouldEnable && !isCurrentComplete);
        }
    }
    
    private void jComboBox2ItemStateChanged(java.awt.event.ItemEvent evt) {
        updateComboBoxStates();
        selected2 = comboBoxes[1].getSelectedIndex() > 0;
    }
    
    private void jComboBox3ItemStateChanged(java.awt.event.ItemEvent evt) {
        updateComboBoxStates();
        selected3 = comboBoxes[2].getSelectedIndex() > 0;
    }
    
    private void jComboBox4ItemStateChanged(java.awt.event.ItemEvent evt) {
        updateComboBoxStates();
        selected4 = comboBoxes[3].getSelectedIndex() > 0;
    }
    
    private void jComboBox5ItemStateChanged(java.awt.event.ItemEvent evt) {
        updateComboBoxStates();
        selected5 = comboBoxes[4].getSelectedIndex() > 0;
    }
    
    private void jComboBox6ItemStateChanged(java.awt.event.ItemEvent evt) {
        updateComboBoxStates();
        selected6 = comboBoxes[5].getSelectedIndex() > 0;
    }
    
    private void jComboBox7ItemStateChanged(java.awt.event.ItemEvent evt) {
        updateComboBoxStates();
        selected7 = comboBoxes[6].getSelectedIndex() > 0;
    }
    
    private void jComboBox8ItemStateChanged(java.awt.event.ItemEvent evt) {
        updateComboBoxStates();
        selected8 = comboBoxes[7].getSelectedIndex() > 0;
    }
    
    private void jComboBox9ItemStateChanged(java.awt.event.ItemEvent evt) {
        updateComboBoxStates();
        selected9 = comboBoxes[8].getSelectedIndex() > 0;
    }
    
    private void jComboBox10ItemStateChanged(java.awt.event.ItemEvent evt) {
        updateComboBoxStates();
        selected10 = comboBoxes[9].getSelectedIndex() > 0;
    }


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
        complete(1);
    }

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



    private void CompleteButton3ActionPerformed(java.awt.event.ActionEvent evt) {
        complete(2);
    }

    private void CompleteButton5ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_CompleteButton5ActionPerformed
        complete(4);
    }

    private void CompleteButton6ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_CompleteButton6ActionPerformed
        complete(5);
    }

    private void CompleteButton7ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_CompleteButton7ActionPerformed
        complete(6);
    }

    private void CompleteButton8ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_CompleteButton8ActionPerformed
        complete(7);
    }

    private void CompleteButton9ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_CompleteButton9ActionPerformed
        complete(8);

    }

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
