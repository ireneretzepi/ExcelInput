package kitreturn;

import java.awt.*;
import java.awt.event.*;
import java.io.*;
import java.time.LocalDate;
import java.util.*;
import java.util.logging.*;
import javax.swing.*;
import javax.swing.filechooser.FileSystemView;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class ExcelInput extends JPanel {

    private static final int MAX_FILES = 10;

    // GUI Components
    private final JComboBox<String>[] comboBoxes = new JComboBox[MAX_FILES];
    private final JButton[] upButtons = new JButton[MAX_FILES];
    private final JButton[] downButtons = new JButton[MAX_FILES];
    private final JButton[] completeButtons = new JButton[MAX_FILES - 1];
    private final JButton[] viewButtons = new JButton[MAX_FILES];

    // Data
    private final ArrayList<String> fileNames = new ArrayList<>();
    private final ArrayList<String> filePaths = new ArrayList<>();
    private final ArrayList<String> viewFiles = new ArrayList<>();
    private final ArrayList<String> completeButtonsPressed = new ArrayList<>();

    private final boolean[] completed = new boolean[MAX_FILES - 1];
    private boolean start = true;
    private String currentFilePath = "";
    private final LocalDate time = LocalDate.now();

    public ExcelInput() {
        initComponents();
        initComboBoxListeners();
        updateButtonStates();
    }

    // GUI initialization
    private void initComponents() {
        setOpaque(false);
        setPreferredSize(new Dimension(1000, 800));

        JLabel headerLabel = new JLabel("Input the Excel files in the order the projects will be completed!");
        JButton addExcelButton = new JButton("Add Excel File");
        JButton removeExcelButton = new JButton("Remove Excel File");
        JButton saveButton = new JButton("Save");

        addExcelButton.addActionListener(e -> addExcel());
        removeExcelButton.addActionListener(e -> removeExcel());
        saveButton.addActionListener(e -> {
            try { saveFile(); } catch (IOException ex) { ex.printStackTrace(); }
        });

        JPanel topPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        topPanel.add(headerLabel);
        topPanel.add(addExcelButton);

        JPanel mainPanel = new JPanel();
        mainPanel.setLayout(new BoxLayout(mainPanel, BoxLayout.Y_AXIS));

        for (int i = 0; i < MAX_FILES; i++) {
            comboBoxes[i] = new JComboBox<>(new String[]{"Select Excel File"});
            upButtons[i] = new JButton("↑");
            downButtons[i] = new JButton("↓");
            viewButtons[i] = new JButton("View👁");

            if (i < MAX_FILES - 1) {
                completeButtons[i] = new JButton("Complete");
                final int idx = i;
                completeButtons[i].addActionListener(e -> complete(idx));
            }

            JPanel row = new JPanel(new FlowLayout(FlowLayout.LEFT));
            row.add(new JLabel((i + 1) + "."));
            row.add(comboBoxes[i]);
            row.add(upButtons[i]);
            row.add(downButtons[i]);
            if (i < MAX_FILES - 1) row.add(completeButtons[i]);
            row.add(viewButtons[i]);
            mainPanel.add(row);

            // Up/Down button actions
            if (i > 0) {
                upButtons[i].addActionListener(e -> swapComboBoxSelection(comboBoxes[i], comboBoxes[i - 1]));
            }
            if (i < MAX_FILES - 1) {
                downButtons[i].addActionListener(e -> swapComboBoxSelection(comboBoxes[i], comboBoxes[i + 1]));
            }
            // View button action
            final int viewIdx = i;
            viewButtons[i].addActionListener(e -> viewFileFromComboBox(comboBoxes[viewIdx]));
        }

        JPanel bottomPanel = new JPanel();
        bottomPanel.add(removeExcelButton);
        bottomPanel.add(saveButton);

        setLayout(new BorderLayout());
        add(topPanel, BorderLayout.NORTH);
        add(mainPanel, BorderLayout.CENTER);
        add(bottomPanel, BorderLayout.SOUTH);
    }

    private void initComboBoxListeners() {
        for (int i = 0; i < comboBoxes.length; i++) {
            final int idx = i;
            comboBoxes[i].addActionListener(e -> updateButtonStates());
        }
    }

    private void addExcel() {
        if (start) {
            resetComboBoxes();
            start = false;
        }
        File selectedFile = promptForExcelFile();
        if (selectedFile == null) return;

        String filename = selectedFile.getName();
        if (fileNames.contains(filename)) {
            JOptionPane.showMessageDialog(this, "The file you chose already exists!");
            return;
        }

        filePaths.add(selectedFile.getAbsolutePath());
        fileNames.add(filename);
        for (JComboBox<String> comboBox : comboBoxes) {
            comboBox.addItem(filename);
        }
        updateButtonStates();
    }

    private void removeExcel() {
        if (fileNames.isEmpty()) return;
        String[] options = fileNames.toArray(new String[0]);
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
        int index = fileNames.indexOf(input);
        if (index != -1) {
            fileNames.remove(index);
            filePaths.remove(index);
            for (JComboBox<String> comboBox : comboBoxes) {
                for (int i = 1; i < comboBox.getItemCount(); i++) {
                    if (comboBox.getItemAt(i).equals(input)) {
                        comboBox.removeItemAt(i);
                        break;
                    }
                }
            }
        }
        updateButtonStates();
    }

    private File promptForExcelFile() {
        JFileChooser chooser = new JFileChooser();
        chooser.setCurrentDirectory(new File(System.getProperty("user.home")));
        int returnValue = chooser.showOpenDialog(this);
        if (returnValue != JFileChooser.APPROVE_OPTION) return null;
        File selectedFile = chooser.getSelectedFile();
        String extension = getFileExtension(selectedFile.getName());
        if (!"xlsx".equalsIgnoreCase(extension)) {
            JOptionPane.showMessageDialog(this, "Please select a valid Excel (.xlsx) file.");
            return null;
        }
        return selectedFile;
    }

    private void resetComboBoxes() {
        for (JComboBox<String> comboBox : comboBoxes) {
            comboBox.setSelectedIndex(0);
        }
    }

    private String getFileExtension(String filename) {
        int dotIndex = filename.lastIndexOf('.');
        return (dotIndex != -1 && dotIndex < filename.length() - 1) ? filename.substring(dotIndex + 1) : "";
    }

    private void swapComboBoxSelection(JComboBox<String> boxA, JComboBox<String> boxB) {
        int indexA = boxA.getSelectedIndex();
        int indexB = boxB.getSelectedIndex();
        boxA.setSelectedIndex(indexB);
        boxB.setSelectedIndex(indexA);
    }

    private void complete(int index) {
        completeButtonsPressed.add("CompleteButton" + (index + 1));
        completed[index] = true;
        comboBoxes[index].setBackground(Color.RED);
        comboBoxes[index].setEnabled(false);
        upButtons[index].setEnabled(false);
        downButtons[index].setEnabled(false);
        viewButtons[index].setEnabled(true);

        String selectedItem = (String) comboBoxes[index].getSelectedItem();
        int fileIndex = fileNames.indexOf(selectedItem);
        if (fileIndex != -1) {
            String outputPath = filePaths.get(fileIndex) + time + ".xlsx";
            viewFiles.add(outputPath);
            // addKitReturn(fileIndex); // Implement as needed
        }
        updateButtonStates();
    }

    private void updateButtonStates() {
        for (int i = 0; i < completeButtons.length; i++) {
            boolean enable = comboBoxes[i].getSelectedIndex() > 0 && comboBoxes[i + 1].getSelectedIndex() > 0 && !completed[i];
            completeButtons[i].setEnabled(enable);
        }
    }

    private void viewFileFromComboBox(JComboBox<String> comboBox) {
        String selected = (String) comboBox.getSelectedItem();
        if (selected == null) return;
        int index = fileNames.indexOf(selected);
        if (index == -1) return;
        String fullPath = filePaths.get(index) + time + ".xlsx";
        try {
            Runtime.getRuntime().exec("excel \"" + fullPath + "\"");
        } catch (IOException ex) {
            Logger.getLogger(ExcelInput.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private void saveFile() throws IOException {
        JFileChooser chooser = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
        int returnValue = chooser.showSaveDialog(null);
        if (returnValue != JFileChooser.APPROVE_OPTION) return;
        File file = chooser.getSelectedFile();
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(file))) {
            writer.write("FileName ArrayList:\n");
            for (String name : fileNames) writer.write(name + "\n");
            writer.write("FilePath ArrayList:\n");
            for (String path : filePaths) writer.write(path + "\n");
            // Add more as needed...
        }
    }
}
