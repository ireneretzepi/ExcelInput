# ExcelInput

## Overview

**ExcelInput** is a specialized application developed at SoftIron to streamline and optimize the process of setting up feeders in machines that assemble hyperdrive components. Its main goal is to track and compare component usage across multiple projects, making it easier to identify reusable feeders and avoid redundant storage and setup.

Previously, feeders were returned to storage after each project with no record of which components were common across future projects. ExcelInput solves this by allowing users to compare Manufacturer Part Numbers (MPNs) from multiple project files, making re-kitting and resource management more efficient.

---

## Features

- **Add Excel File:**  
  Easily add Excel files containing feeder/component information. The application stores both the file name (`FileName`) and its path (`FilePath`).

- **Remove Excel File:**  
  Remove any previously added Excel file using a dropdown list of all currently loaded files.

- **Select Excel File:**  
  Choose from a dropdown of added files to select the "base" file for comparison.

- **Multi-File Comparison:**  
  Select at least two files (the base file and one or more "next" files) to enable the **Complete** button. The application compares MPNs between the selected files.

- **Comparison Output:**  
  On pressing **Complete**, a new Excel file is generated. This file is a copy of the selected base file, with an added column named **Re-kit**. Each column in the new file shows the common MPNs between the base file and each of the selected next files.

- **View Result:**  
  After comparison, the **View** button is enabled, allowing users to view the newly created Excel file.

- **Reorder Selection:**  
  Use **Up** and **Down** buttons to change the sequence of selected files. These buttons are disabled after comparison is complete.

- **State Saving:**  
  Save the current state of the application and reopen it later to continue work without losing progress.

- **Help & Guidelines:**  
  A **Help** button provides users with detailed guidelines on how to use the application.

---

## How It Works

1. **Add Project Files:**  
   Click **Add Excel File** to import files containing feeder/component information.

2. **Select Files for Comparison:**  
   Use the **Select Excel File** dropdown to choose the base file. Then, select at least one additional file for comparison. The **Complete** button becomes active when at least two files are selected.

3. **Run Comparison:**  
   Click **Complete** to generate a new Excel file highlighting common Manufacturer Part Numbers across the selected files. The new file will have a **Re-kit** column for each comparison.

4. **View Results:**  
   Use the **View** button to open and review the generated Excel file.

5. **Manage Files:**  
   Remove files or reorder the selection as needed. Note: After running a comparison, the reorder buttons are disabled for the selected files.

6. **Save/Load State:**  
   Save your progress at any time and reload it later.

7. **Get Help:**  
   Click the **Help** button for step-by-step instructions and troubleshooting tips.

---


## Usage Example

1. Add Excel files for Project A, Project B, and Project C.  
2. Select Project A as the base file.  
3. Select Project B and Project C as comparison files.  
4. Click **Complete** to generate a new Excel file showing common MPNs.  
5. Click **View** to review the results.  
6. Save your session to continue later.

---

## Support

If you have questions or need assistance, click the **Help** button in the application for detailed guidance.

---

**ExcelInput** makes feeder management and component tracking easy, efficient, and transparent—helping you make the most of your resources across all your projects!
