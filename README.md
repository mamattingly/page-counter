# Document Page Counter Script

This PowerShell script is designed to count the number of pages in Word and PDF documents within a specified folder. It supports `.docx` and `.pdf` file formats and provides options for writing the results to a CSV file, including a summary of the counts, and generating a CSV file with the current date.

## Features

- **Counts Pages**: Counts the number of pages in `.docx` and `.pdf` files.
- **CSV Output**: Optionally writes the results to a CSV file. You can choose to overwrite or append to an existing CSV file.
- **Date in Filename**: Optionally includes the current date in the CSV file name.
- **Summary Information**: Optionally displays a summary of total pages and files processed.

## Usage

1. **Prepare Your Environment**:

   - Unzip the file to any directory of your choice.
   - Change the file extensions on StartCount.txt and DocumentPageCounter.txt to StartCount.bat and DocumentPageCounter.ps1
   - Move `.docx` and `.pdf` files to the "Documents" folder in the unzipped directory. (*If this is the first time running the script, there will be test files that should be deleted.*) - this is optional.
   - From here you can alter the parameters in DocumentPageCounter.ps1 to your liking and run the program with StartCount.bat or continue below.
2. **Run the Script**:

   - Open PowerShell.
   - Navigate to the directory where the script is located.
   - Execute the script with desired parameters. For example:

     ```powershell
     Get-FolderPageCounts -writeCSV $true -folderPath "C:\Path\To\Documents" -includeDateInFileName $true -includeSummary $true
     ```
   - Parameters:

     - `-folderPath`: (Optional) Path to the folder containing documents. If omitted, defaults to "./Documents".
     - `-writeCSV`: (Optional) Flag to enable or disable CSV writing. Defaults to `false`.
     - `-includeDateInFileName`: (Optional) Flag to include the current date in the CSV file name. Defaults to `false`.
     - `-includeSummary`: (Optional) Flag to display summary information. Defaults to `false`.
3. **Review Results**:

   - The script will output the file name, page count, and document type for each `.docx` and `.pdf` file in the specified folder.
   - If CSV writing is enabled, the results will be saved to a CSV file in the Downloads folder.

## Requirements

- PowerShell
- Microsoft Word installed on the machine (for counting pages in Word documents)
- Permissions to access the files in the specified folder

## Author

* **Michael Mattingly**

  GitHub: [mamattingly](https://github.com/mamattingly)

## Badges

![MIT License](https://img.shields.io/badge/License-MIT-yellow.svg)

![PowerShell](https://img.shields.io/badge/PowerShell-1f425f.svg?style=flat&logo=powershell&logoColor=white)

## License

This script is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.
