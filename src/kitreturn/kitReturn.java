package kitreturn;

import javax.swing.*;
import javax.swing.filechooser.FileSystemView;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.*;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;
import java.util.logging.Level;
import java.util.logging.Logger;

public class KitReturnWindow extends JFrame {
    private final ExcelInput excelInput = new ExcelInput();

    private final List<String> fileNameList = new ArrayList<>();
    private final List<String> filePathList = new ArrayList<>();
    private final List<String> selectedComboBoxes = new ArrayList<>();
    private final List<String> viewFiles = new ArrayList<>();
    private final List<String> completeButtonsPressed = new ArrayList<>();
    private final List<String> rawFileLines = new ArrayList<>();
    private String disableUpDown = null;

    private final JPanel mainPanel = new JPanel();

    public KitReturnWindow() {
        setupUI();
    }

    private void setupUI() {
        setTitle("Kit Return");
        setSize(800, 600);
        setLocationRelativeTo(null);
        setDefaultCloseOperation(DISPOSE_ON_CLOSE);

        JMenuBar menuBar = new JMenuBar();
        JMenu fileMenu = new JMenu("File");
        JMenu helpMenu = new JMenu("Help");

        JCheckBoxMenuItem openFileItem = new JCheckBoxMenuItem("Open File");
        openFileItem.addActionListener(this::handleOpenFile);
        fileMenu.add(openFileItem);

        JCheckBoxMenuItem helpItem = new JCheckBoxMenuItem("Help");
        helpItem.addActionListener(this::showHelp);
        helpMenu.add(helpItem);

        menuBar.add(fileMenu);
        menuBar.add(helpMenu);
        setJMenuBar(menuBar);

        setContentPane(mainPanel);
        setVisible(true);
    }

    private void showHelp(ActionEvent evt) {
        String helpText = "<html><h1 align='center'>Help</h1>"
            + "<h2>How to function the application?</h2>"
            + "<p>Press the “Add Excel File” button to add files to the drop boxes and add "
            + "projects to the list of existing projects. If you accidentally add the same "
            + "file twice or open a non-Excel file, error messages will be displayed.</p>"
            + "<p>Use the “Select Excel File” drop boxes to select the projects that will be "
            + "completed in the sequence they will be completed. Using the ⇧⇩ buttons you can "
            + "swap the selected files. Once you press the complete button, the project won't be "
            + "editable or swappable with other projects.</p>"
            + "<p>If you add an Excel file by accident, you can remove it by pressing the “Remove Excel File” "
            + "button, which will display a drop-down of all added files so far for selection and removal.</p>"
            + "<p>When you press the “complete” button, 're-kit' as well as the name of the project "
            + "that the manufacturer parts will be re-kitted to will appear in the Excel next to the "
            + "parts. To view the Excel file, press 'view' next to the project. The 'view' button will "
            + "only be enabled once the 'complete' button has been pressed for that project.</p>"
            + "<p>After completing the project series, you can save the status of completed files "
            + "and continue another time by opening that file.</p></html>";
        JOptionPane.showMessageDialog(this, new JLabel(helpText), "Help", JOptionPane.INFORMATION_MESSAGE);
    }

    private void handleOpenFile(ActionEvent evt) {
        JFileChooser fileChooser = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
        fileChooser.setCurrentDirectory(new File(System.getProperty("user.home")));

        int result = fileChooser.showOpenDialog(this);
        if (result != JFileChooser.APPROVE_OPTION) return;

        File selectedFile = fileChooser.getSelectedFile();
        while (!selectedFile.getName().endsWith(".txt")) {
            JOptionPane.showMessageDialog(this, "Please select a .txt file!");
            result = fileChooser.showOpenDialog(this);
            if (result != JFileChooser.APPROVE_OPTION) return;
            selectedFile = fileChooser.getSelectedFile();
        }

        readAndParseFile(selectedFile);
    }

    private void readAndParseFile(File file) {
        try (Scanner scanner = new Scanner(file)) {
            while (scanner.hasNextLine()) {
                rawFileLines.add(scanner.nextLine().trim());
            }
            parseSections(rawFileLines);
            excelInput.readFile(fileNameList, filePathList, selectedComboBoxes, disableUpDown, viewFiles, completeButtonsPressed);
        } catch (IOException ex) {
            Logger.getLogger(KitReturnWindow.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private void parseSections(List<String> lines) {
        for (int i = 0; i < lines.size(); i++) {
            String line = lines.get(i);

            if (line.equalsIgnoreCase("FileName ArrayList:")) {
                i = collectUntil(lines, i, "FilePath ArrayList:", fileNameList);
            } else if (line.equalsIgnoreCase("FilePath ArrayList:")) {
                i = collectUntil(lines, i, "Selected Combo Boxes:", filePathList);
            } else if (line.equalsIgnoreCase("Selected Combo Boxes:")) {
                i = collectUntil(lines, i, "Number of up/down buttons that should be disabled:", selectedComboBoxes);
            } else if (line.contains("Number of up/down buttons that should be disabled:")) {
                disableUpDown = line.split(":")[1].trim();
            } else if (line.equalsIgnoreCase("View files created:")) {
                i = collectUntil(lines, i, "Complete Buttons Pressed:", viewFiles);
            } else if (line.equalsIgnoreCase("Complete Buttons Pressed:")) {
                for (int j = i + 1; j < lines.size(); j++) {
                    completeButtonsPressed.add(lines.get(j));
                }
                break;
            }
        }
    }

    private int collectUntil(List<String> lines, int index, String stopMarker, List<String> target) {
        int i = index;
        while (i + 1 < lines.size() && !lines.get(i + 1).equalsIgnoreCase(stopMarker)) {
            i++;
            target.add(lines.get(i));
        }
        return i;
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(KitReturnWindow::new);
    }
}

