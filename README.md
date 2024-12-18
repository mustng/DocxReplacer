# Word Document Placeholder Replacement Tool (DocxReplacer)

DocxReplacer is a lightweight tool designed to automate the replacement of placeholders in Microsoft Word `.docx` templates. While tailored for generating leases, it is versatile enough for contracts, reports, personalized letters, and other dynamic documents. It supports number-to-words conversion and customizable date formatting.

---

## Features

- Replace placeholders in `.docx` templates with values specified in a configuration file.
- Automatically convert numeric values into words (e.g., `6` becomes `six` using `[LeaseTermAmountToWord]`).
- Handle date placeholders with customizable formatting and ordinal suffixes (e.g., `November 21st, 2024`).
- Create dynamically named output files based on content and timestamp.
- Simple configuration through a plain text file (`config.txt`).

---

## Download

Outright download [DocxReplacer.zip](https://github.com/user-attachments/files/18174354/DocxReplacer.zip)

For users who don’t want to build the project themselves, download the pre-built executable:

- [Download DocxReplacer Executable](https://github.com/mustng/DocxReplacer/releases)

---

## Issues

**Known Issue**: If the `.docx` file contains bold placeholders, it may cause the entire paragraph to appear bold. Use consistent formatting within templates to avoid this issue.

### Instructions for Non-Programmers

#### Prerequisites
1. **Install .NET Runtime**:
   - Download and install the [.NET 8.0 Runtime](https://dotnet.microsoft.com/download/dotnet/8.0) for your operating system:
     - For **Windows**, download the installer and follow the on-screen instructions.
     - For **Mac**, download the `.pkg` installer and follow the installation steps.
   - Verify the installation by running the following command in your terminal or command prompt:
     ```bash
     dotnet --version
     ```
     You should see the installed version of .NET (e.g., `8.0.x`).

---

#### How to Run on **Windows**
1. **Download the Tool**:
   - Go to the [Releases](https://github.com/mustng/DocxReplacer/releases) page and download the latest `.zip` file.
   - Extract the contents of the `.zip` file to a folder on your computer.

2. **Prepare Your Files**:
   - Place your Word `.docx` template and `config.txt` file in the same folder as the `WordDocumentPlaceholderReplacementTool.exe` file.

3. **Run the Tool**:
   - Open the folder where the `.exe` file is located.
   - Double-click the `WordDocumentPlaceholderReplacementTool.exe` file to run the application.

---

#### How to Run on **Mac**
1. **Download the Tool**:
   - Go to the [Releases](https://github.com/mustng/DocxReplacer/releases) page and download the latest `.zip` file.
   - Extract the contents of the `.zip` file to a folder on your computer.

2. **Prepare Your Files**:
   - Place your Word `.docx` template and `config.txt` file in the same folder as the tool's executable file.

3. **Run the Tool**:
   - Open the Terminal and navigate to the folder containing the tool:
     ```bash
     cd /path/to/extracted/folder
     ```
   - Execute the tool with the following command:
     ```bash
     dotnet WordDocumentPlaceholderReplacementTool.dll
     ```
   - The output file will be saved in the directory specified in `config.txt` or alongside the input `.docx`.

---

#### Notes
- Ensure your `config.txt` file follows the [Configuration](#configuration) guidelines.
- If the `OutputDocxPath` is left empty, the output file will be saved in the same folder as the input `.docx`, with a dynamically generated name.
- On **Mac**, you must run the `.dll` file using the `dotnet` command.

---

## How to Build from Source (for Developers)

1. Clone the repository to your local machine:
   git clone https://github.com/mustng/DocxReplacer.git

2. Navigate to the project directory:
   cd DocxReplacer

---

## How to Run the Tool

1. Open the project in your preferred IDE (e.g., Visual Studio).
2. Ensure you have the necessary dependencies installed:
   - **DocumentFormat.OpenXml**
3. Place your Word `.docx` template and `config.txt` in the same directory as the executable.
4. Update `config.txt` with the required values (see the Configuration section below).
5. Run the application. The output file will be saved in the directory specified in the `OutputDocxPath` or, if left blank, alongside the input file with a dynamically generated name.

---

## Special `ToWord` Action

DocxReplacer includes a special feature for converting numeric placeholders into their word equivalents.

Example: If `LeaseTermAmount = "6"` is specified in `config.txt` and the `.docx` file contains `[LeaseTermAmountToWord]`, it will be replaced with `six`.

This feature allows seamless integration of human-readable numbers into your documents.

---

## Configuration

The tool uses a plain text configuration file (`config.txt`) for input. Below is an example:

InputDocxPath = "D:\\Test\\Residential2024v1.docx"  
OutputDocxPath = "" // Leave empty to use LesseeName with date and time and add to InputDocxPath directory  
Date = "" // Leave blank if you want today's date (e.g., November 21st, 2024)  
LessorName = "Billy Hughes"  
LesseeName = "Breanna Smith"  
PropertyAddress = "123 Tech Street Tech City"  
County = "Duval"  
LeaseTermAmount = "6"  
LeaseTermPeriod = "months"  
LeaseStartDate = "October 17th, 2024"  
LeaseEndDate = "April 1st, 2025"  

---

## Example Word Template

Here’s an example `.docx` template to be used with DocxReplacer:

Residential Lease

This Lease Agreement ("Agreement") is made and entered into on [Date], between [LessorName] (hereinafter referred to as "Lessor") and [LesseeName] (hereinafter collectively referred to as "Lessee"). The Lessor hereby agrees to lease to the Lessee the premises located at [PropertyAddress], situated in [County] County, along with all appurtenances thereto, for a term of [LeaseTermAmountToWord] ([LeaseTermAmount]) [LeaseTermPeriod], commencing on [LeaseStartDate], and ending at 12:00 PM (noon) on [LeaseEndDate]. Thereafter, the lease shall convert to a month-to-month tenancy unless and until the Lessor and Lessee agree upon and execute a new lease agreement.

---

## Output Example

Using the above `config.txt` and Word template, the tool will generate:

Residential Lease

This Lease Agreement ("Agreement") is made and entered into on November 21st, 2024, between Billy Hughes (hereinafter referred to as "Lessor") and Breanna Smith (hereinafter collectively referred to as "Lessee"). The Lessor hereby agrees to lease to the Lessee the premises located at 123 Tech Street Tech City, situated in Duval County, along with all appurtenances thereto, for a term of six (6) months, commencing on October 17th, 2024, and ending at 12:00 PM (noon) on April 1st, 2025. Thereafter, the lease shall convert to a month-to-month tenancy unless and until the Lessor and Lessee agree upon and execute a new lease agreement.

---

## Contributing

Contributions are welcome! Feel free to fork the repository and submit a pull request with your improvements.

---

## License

This project is licensed under the MIT License.
