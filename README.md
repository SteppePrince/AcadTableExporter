# AcadTableExporter
This is a tool to export tables from Acad to CSV files.
# Install
 - Use the NuGet package manager to search for ExportTableToCSV.
 - Alternatively, you can run the following command in the terminal:
```
dotnet add package ExportTableToCSV
```
# Usage
### Loading the DLL in CAD Software
- Open your CAD software.
- Enter the NETLOAD command and press Enter.
- In the dialog that appears, navigate to your ExportTableToCSV.dll file, select it, and confirm.
### Exporting Tables to CSV
- In the CAD command line, enter ExportTableToCSV and press Enter.
- Follow the prompts to select the table object you wish to export.
- Upon confirmation, the library will process the selected table and export it to a CSV file.
# Examples
The example dwg file and the target csv are shown.
# Attention
- Users are required to manually download the provided source code files and integrate them into your own .NET project. Ensure that your project targets the .NET Framework 4.7.2, as this is the framework for which the library was designed.
- The library depends on specific CAD software DLL files that are not included in this repository. Users must obtain these DLL files directly from their CAD software installation or the software provider's website. Please ensure that you have the appropriate licenses and rights to use these DLL files.

