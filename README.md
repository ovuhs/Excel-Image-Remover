# Excel Image Remover 📊

A lightweight Python tool designed to optimize Excel files by automatically removing all embedded images from every sheet within a workbook. This is particularly useful for reducing file size or cleaning up reports before sharing.

## 🚀 Features

- **Bulk Removal:** Scans all sheets in the workbook and deletes every image found.
- **Format Support:** Works with standard Excel files (`.xlsx`) and Macro-enabled files (`.xlsm`).
- **Safe Processing:** Does not overwrite the original file. It saves a new copy named `filename_no_images.xlsx`.
- **User Friendly:** Includes a simple file selection window (GUI).
- **Log Summary:** Shows exactly how many images were removed from which sheets.

## 🛠️ Prerequisites

To run this script, you need Python installed on your system. It uses the `openpyxl` library for Excel manipulation.

## 📦 Installation

1. Clone this repository:
   ```bash
   git clone [https://github.com/ovuhs/Excel-Image-Remover.git](https://github.com/ovuhs/Excel-Image-Remover.git)
   cd Excel-Image-Remover
