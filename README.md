# Crystal Reports Definition Extractor

A simple, powerful command-line utility for extracting key definition information from SAP Crystal Reports (`.rpt`) files. This tool is designed for developers, analysts, and administrators who need to document, analyze, or migrate reports without manually opening each one.

It extracts critical metadata into a human-readable text file, including:
- Data source connection information
- Tables, Views, or Stored Procedures used
- All user-facing parameters
- Record and Group selection formulas
- The text of all custom formulas

This tool is built using C# and the .NET Framework.

## Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation and Setup](#installation-and-setup)
- [Usage](#usage)
- [Example](#example)
- [Building from Source](#building-from-source)
- [Troubleshooting](#troubleshooting)

## Features

- **Data Source Identification**: Automatically detects and lists the server, database, and tables/views or stored procedures a report is connected to.
- **Parameter Extraction**: Lists all parameters with their name, data type, and prompt text.
- **Formula Extraction**: Dumps the complete text of all record selection, group selection, and custom formulas.
- **Grouping Information**: Shows which fields the report is grouped by and their sort direction.
- **Command-Line Interface**: Easy to automate and use in batch scripts to process multiple reports at once.

## Prerequisites

To run this application, the target machine must have the following installed:

1.  **.NET Framework 4.8** (or the version targeted by the project). This is pre-installed on most modern Windows systems.
2.  **SAP Crystal Reports Runtime for .NET (64-bit)**. The application is built for the x64 platform and requires the corresponding 64-bit runtime.

## Installation and Setup

To use the tool on any computer, follow these two steps:

1.  **Install the Crystal Reports Runtime:**
    - Download the **64-bit (x64)** Crystal Reports Runtime installer (`CRRuntime_64bit_13_0_xx.msi`) from the official SAP website.
    - Run the installer on the target machine. This only needs to be done once per machine.

2.  **Deploy the Application:**
    - Unzip the `CrystalReportDocumenter.zip` release file to a convenient location on your computer (e.g., `C:\Tools\ReportDocumenter\`).

## Usage

The application is run from the command line (`cmd.exe` or PowerShell). It requires two arguments: the path to the input `.rpt` file and the path for the output `.txt` definition file.

**Syntax:**

```sh
CrystalReportDocumenter.exe "<input_report_path>" "<output_definition_path>"
```

-   `<input_report_path>`: The full path to the `.rpt` file you want to analyze.
-   `<output_definition_path>`: The full path where you want to save the resulting `.txt` file.

**Important:** Always enclose file paths in double quotes (`"`) to handle any spaces in the names.

## Example

1.  Open a Command Prompt.
2.  Navigate to the directory where you unzipped the application files.

    ```cmd
    cd C:\Tools\ReportDocumenter
    ```

3.  Run the command, pointing to your report file.

    ```cmd
    CrystalReportDocumenter.exe "C:\Users\Thoma\Reports\SalesByRegion.rpt" "C:\Users\Thoma\Desktop\SalesByRegion_Definition.txt"
    ```

4.  After the command finishes, a file named `SalesByRegion_Definition.txt` will be created on your desktop with all the extracted report metadata.

## Building from Source

If you need to modify or rebuild the application, follow these steps:

1.  **Clone the Repository:**
    ```sh
    git clone <repository_url>
    ```
2.  **Install Prerequisites:**
    - Ensure you have Visual Studio installed with .NET Framework development tools.
    - Install **SAP Crystal Reports for Visual Studio** from the official SAP website. This provides the necessary SDK libraries.
3.  **Open the Solution:**
    - Open the `.sln` file in Visual Studio.
4.  **Configure the Build:**
    - Set the solution configuration to **Release**.
    - Set the solution platform to **x64**.
5.  **Build the Project:**
    - Go to **Build -> Rebuild Solution**.
    - The output files will be located in the `bin\x64\Release` directory.

## Troubleshooting

-   **"File Not Found" or "Unable to load DLL" Error:** This almost always means the **SAP Crystal Reports Runtime for .NET (64-bit)** is not installed on the machine, or the wrong version (32-bit) was installed. Ensure the 64-bit runtime is installed correctly.
-   **Application Fails Silently:** Make sure you have the correct version of the .NET Framework installed.

---
```
