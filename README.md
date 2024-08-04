# Document Page Counter Script

This PowerShell script is designed to count the number of pages in Word and PDF documents within a specified folder. It supports `.docx` and `.pdf` file formats and displays the page count for each document in the console.

## Usage

1. Save the script to a `.ps1` file (e.g., `DocumentPageCount.ps1`).
2. Modify the `$folderPath` variable to point to the directory containing your documents.
3. Place `.pdf` or `.docx` documents in the folder that the path is pointing to.
4. Run the script.

The script will output the file name, page count, and document type for each `.docx` and `.pdf` file in the specified folder.

## Requirements

* PowerShell
* Microsoft Word installed on the machine (for counting pages in Word documents)
* Ensure the script has permissions to access the files in the specified folder.

## Author

* **Michael Mattingly**

  GitHub: [mamattingly](https://github.com/mamattingly)

## Badges

![MIT License](https://img.shields.io/badge/License-MIT-yellow.svg)

![PowerShell](https://img.shields.io/badge/PowerShell-1f425f.svg?style=flat&logo=powershell&logoColor=white)

## License

This script is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.
