# MergePDF Project

## Overview
The MergePDF project is a PowerShell-based application designed to merge multiple PDF files into a single document. It provides a user-friendly interface that allows users to drag and drop PDF files, reorder them, and merge them with ease.

## Project Structure
```
MergePDF
├── src
│   ├── MergePDF.ps1       # Main PowerShell script for merging PDF files
│   └── ui
│       └── MergePDF.xaml   # XAML definition for the user interface
├── README.md               # Documentation for the project
└── .gitignore              # Git ignore file
```

## Files Description

### `src/MergePDF.ps1`
This file contains the main logic for the application. It handles:
- Loading the user interface defined in the XAML file.
- Implementing drag-and-drop functionality for adding PDF files.
- Merging the selected PDF files using the PSWritePDF module.

### `src/ui/MergePDF.xaml`
This file defines the layout and visual elements of the application window. It includes:
- Buttons for merging and deleting selected files.
- A list box for displaying the added PDF files with their indices.

## Usage Instructions
1. Ensure that you have PowerShell installed on your system.
2. Install the PSWritePDF module if it is not already installed. The script will prompt for installation if necessary.
3. Run the `MergePDF.ps1` script in PowerShell.
4. Drag and drop PDF files into the application window.
5. Reorder the files as needed.
6. Click the "Merge" button to combine the selected PDF files into a single document.

## Setup Instructions
- Clone the repository to your local machine.
- Open PowerShell and navigate to the `src` directory.
- Execute the `MergePDF.ps1` script to launch the application.

## Git Ignore
The `.gitignore` file is included to prevent unnecessary files from being tracked by Git, ensuring a clean repository.