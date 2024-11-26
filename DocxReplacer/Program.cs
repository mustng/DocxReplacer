using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Xceed.Document.NET;
using Xceed.Words.NET;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Word Document Placeholder Replacement Tool");

        // Get the full path to the configuration file in the application's directory
        string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
        string configFilePath = Path.Combine(appDirectory, "config.txt");

        if (!File.Exists(configFilePath))
        {
            Console.WriteLine("The config.txt file is missing in the application's directory. Exiting.");
            return;
        }

        // Declare replacements and file paths
        Dictionary<string, string> replacements;
        string inputDocxPath;
        string outputDocxPath = string.Empty;

        try
        {
            replacements = ParseConfigFile(configFilePath);

            // Validate required fields
            if (!replacements.ContainsKey("InputDocxPath") || string.IsNullOrWhiteSpace(replacements["InputDocxPath"]))
            {
                Console.WriteLine("InputDocxPath is missing or invalid in the config file. Exiting.");
                return;
            }

            if (!replacements.ContainsKey("LesseeName") || string.IsNullOrWhiteSpace(replacements["LesseeName"]))
            {
                Console.WriteLine("LesseeName is required and must be present in the config file. Exiting.");
                return;
            }

            // Extract paths
            inputDocxPath = replacements["InputDocxPath"];
            if (!File.Exists(inputDocxPath) || Path.GetExtension(inputDocxPath) != ".docx")
            {
                Console.WriteLine("The specified InputDocxPath is invalid or the file does not exist. Exiting.");
                return;
            }

            if (replacements.TryGetValue("OutputDocxPath", out var customOutputPath) && !string.IsNullOrWhiteSpace(customOutputPath))
            {
                outputDocxPath = customOutputPath;
            }
            else
            {
                // Generate a file name using LesseeName and the current date/time
                if (!replacements.TryGetValue("LesseeName", out var lesseeName) || string.IsNullOrWhiteSpace(lesseeName))
                {
                    Console.WriteLine("LesseeName is required and must be present in the config file. Exiting.");
                    return;
                }

                string sanitizedFileName = $"{lesseeName.Replace(" ", "")}_{DateTime.Now:yyyyMMddHHmmss}.docx";
                outputDocxPath = Path.Combine(Path.GetDirectoryName(inputDocxPath), sanitizedFileName);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error reading configuration: {ex.Message}");
            return;
        }

        try
        {
            // Perform replacements and save the new Word document
            ReplacePlaceholdersInDocx(inputDocxPath, replacements, outputDocxPath);

            Console.WriteLine($"Replacement completed. Updated file saved at: {outputDocxPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing the Word document: {ex.Message}");
        }
    }

    // Method to parse the config file
    static Dictionary<string, string> ParseConfigFile(string filePath)
    {
        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var line in File.ReadAllLines(filePath))
        {
            if (string.IsNullOrWhiteSpace(line) || line.StartsWith("//")) continue; // Skip comments and empty lines
            var parts = line.Split(new[] { '=' }, 2);
            if (parts.Length == 2)
            {
                var key = parts[0].Trim();
                var value = parts[1].Split(new[] { "//" }, StringSplitOptions.None)[0].Trim().Trim('"'); // Get value before comment, trim quotes
                dict[key] = value;
            }
        }
        return dict;
    }

    // Method to replace placeholders in the Word document
    static void ReplacePlaceholdersInDocx(string inputFilePath, Dictionary<string, string> replacements, string outputDocxPath)
    {
        try
        {
            // Ensure the output directory exists
            string outputDirectory = Path.GetDirectoryName(outputDocxPath);
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
                Console.WriteLine($"Created directory: {outputDirectory}");
            }

            // Remove existing file if necessary
            if (File.Exists(outputDocxPath))
            {
                File.Delete(outputDocxPath);
                Console.WriteLine($"Deleted existing file: {outputDocxPath}");
            }

            // Load the Word document
            using (var document = DocX.Load(inputFilePath))
            {
                // Get all placeholders in the document
                var placeholders = GetPlaceholdersFromDocument(document);

                // Perform replacements
                foreach (var placeholder in placeholders)
                {
                    Console.WriteLine($"Processing placeholder: {placeholder}");

                    if (placeholder.EndsWith("ToWord"))
                    {
                        // Handle "ToWord" placeholders dynamically
                        string baseKey = placeholder.Replace("ToWord", ""); // Extract base key
                        if (replacements.TryGetValue(baseKey, out string numericValue) &&
                            int.TryParse(numericValue, out int numeric))
                        {
                            string wordValue = ConvertNumberToWords(numeric);
                            ReplaceTextWithObject(document, $"[{placeholder}]", wordValue);
                            Console.WriteLine($"Replaced '[{placeholder}]' with '{wordValue}'");
                            continue;
                        }
                        else
                        {
                            Console.WriteLine($"Error: Unable to resolve or convert value for '{baseKey}'.");
                        }
                    }
                    else if (placeholder == "Date" && string.IsNullOrWhiteSpace(replacements[placeholder]))
                    {
                        // Handle date formatting with ordinal suffix
                        DateTime today = DateTime.Now;
                        int day = today.Day;
                        string formattedDate = $"{today:MMMM} {day}{GetOrdinalSuffix(day)}, {today.Year}";
                        ReplaceTextWithObject(document, $"[{placeholder}]", formattedDate);
                        Console.WriteLine($"Replaced '[{placeholder}]' with '{formattedDate}'");
                    }
                    else if (replacements.TryGetValue(placeholder, out string replacementValue))
                    {
                        // Handle regular placeholders
                        ReplaceTextWithObject(document, $"[{placeholder}]", replacementValue);
                        Console.WriteLine($"Replaced '[{placeholder}]' with '{replacementValue}'");
                    }
                    else
                    {
                        Console.WriteLine($"Warning: Placeholder '[{placeholder}]' not found in replacements.");
                    }
                }

                // Save the updated document
                Console.WriteLine($"Attempting to save file at: {outputDocxPath}");
                document.SaveAs(outputDocxPath);
                Console.WriteLine($"File saved successfully at {outputDocxPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during processing: {ex.Message}");
        }
    }
    static void ReplaceTextWithObject(DocX document, string placeholder, string replacementText)
    {
        foreach (var paragraph in document.Paragraphs)
        {
            // Check if the paragraph contains the placeholder
            if (paragraph.Text.Contains(placeholder))
            {
                // Replace the placeholder with the replacement text
                paragraph.ReplaceText(placeholder, replacementText); // This works for string replacement
                Console.WriteLine($"Replaced '{placeholder}' with '{replacementText}' in paragraph.");
            }
        }
    }

    // Helper to extract all placeholders from the document
    static List<string> GetPlaceholdersFromDocument(DocX document)
    {
        var placeholders = new List<string>();

        // Search for placeholders enclosed in square brackets
        foreach (var paragraph in document.Paragraphs)
        {
            var matches = Regex.Matches(paragraph.Text, @"\[(.*?)\]");
            foreach (Match match in matches)
            {
                placeholders.Add(match.Groups[1].Value); // Extract placeholder name without brackets
            }
        }

        return placeholders;
    }

    // Helper to convert numbers to words
    static string ConvertNumberToWords(int number)
    {
        if (number == 0) return "zero";

        string[] units = { "", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen" };
        string[] tens = { "", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety" };

        if (number < 20)
            return units[number];

        if (number < 100)
            return tens[number / 10] + (number % 10 > 0 ? "-" + units[number % 10] : "");

        if (number < 1000)
            return units[number / 100] + " hundred" + (number % 100 > 0 ? " and " + ConvertNumberToWords(number % 100) : "");

        if (number < 1000000)
            return ConvertNumberToWords(number / 1000) + " thousand" + (number % 1000 > 0 ? " " + ConvertNumberToWords(number % 1000) : "");

        if (number < 1000000000)
            return ConvertNumberToWords(number / 1000000) + " million" + (number % 1000000 > 0 ? " " + ConvertNumberToWords(number % 1000000) : "");

        return ConvertNumberToWords(number / 1000000000) + " billion" + (number % 1000000000 > 0 ? " " + ConvertNumberToWords(number % 1000000000) : "");
    }

    // Helper to add ordinal suffix to dates
    static string GetOrdinalSuffix(int day)
    {
        if (day % 100 >= 11 && day % 100 <= 13) return "th"; // Handle special cases: 11th, 12th, 13th
        return (day % 10) switch
        {
            1 => "st",
            2 => "nd",
            3 => "rd",
            _ => "th"
        };
    }
}
