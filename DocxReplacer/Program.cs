using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

Console.WriteLine("Word Document Placeholder Replacement Tool");

// Get the full path to the configuration file in the application's directory
string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
string configFilePath = Path.Combine(appDirectory, "config.txt");

if (!File.Exists(configFilePath))
{
    Console.WriteLine("The config.txt file is missing in the application's directory. Exiting.");
    return;
}

// Load the configuration from the text file
Dictionary<string, string> replacements;
string inputDocxPath;
string outputDocxPath;

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
    // Perform replacements and create a new Word file
    ReplacePlaceholdersInDocx(inputDocxPath, replacements, outputDocxPath);

    Console.WriteLine($"Replacement completed. Updated file saved at: {outputDocxPath}");
}
catch (Exception ex)
{
    Console.WriteLine($"Error processing the Word document: {ex.Message}");
}

// Method to parse the config file
Dictionary<string, string> ParseConfigFile(string filePath)
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
void ReplacePlaceholdersInDocx(string inputFilePath, Dictionary<string, string> replacements, string outputFilePath)
{
    // Copy the input file to the output file
    File.Copy(inputFilePath, outputFilePath, true);

    // Open the output file for editing
    using (var wordDoc = WordprocessingDocument.Open(outputFilePath, true))
    {
        var body = wordDoc.MainDocumentPart?.Document.Body;
        if (body == null)
            throw new Exception("Invalid Word document structure.");

        // Process each paragraph
        foreach (var paragraph in body.Elements<Paragraph>())
        {
            var textElements = paragraph.Descendants<Text>().ToList();
            if (!textElements.Any()) continue;

            // Combine all text runs in the paragraph
            string combinedText = string.Join("", textElements.Select(t => t.Text));
            Console.WriteLine($"Original Paragraph Text: {combinedText}");

            // Find and replace placeholders
            foreach (var key in replacements.Keys)
            {
                string placeholder = $"[{key}]";
                string toWordPlaceholder = $"[{key}ToWord]";

                // Replace "ToWord" placeholders dynamically
                if (combinedText.Contains(toWordPlaceholder))
                {
                    if (int.TryParse(replacements[key], out int numericValue))
                    {
                        string wordValue = ConvertNumberToWords(numericValue);
                        combinedText = combinedText.Replace(toWordPlaceholder, wordValue);
                        Console.WriteLine($"Replaced '{toWordPlaceholder}' with '{wordValue}'");
                    }
                    else
                    {
                        Console.WriteLine($"Error: Value for '{key}' is not a valid number.");
                    }
                }

                // Replace regular placeholders
                if (combinedText.Contains(placeholder))
                {
                    if (key == "Date" && string.IsNullOrWhiteSpace(replacements[key]))
                    {
                        // Handle date formatting with ordinal suffix
                        DateTime today = DateTime.Now;
                        int day = today.Day;
                        string formattedDate = $"{today:MMMM} {day}{GetOrdinalSuffix(day)}, {today.Year}";
                        combinedText = combinedText.Replace(placeholder, formattedDate);
                        Console.WriteLine($"Replaced '[{key}]' with '{formattedDate}'");
                    }
                    else
                    {
                        combinedText = combinedText.Replace(placeholder, replacements[key]);
                        Console.WriteLine($"Replaced '[{key}]' with '{replacements[key]}'");
                    }
                }
            }

            // Update the text runs with the modified content
            for (int i = 0; i < textElements.Count; i++)
            {
                textElements[i].Text = i == 0 ? combinedText : string.Empty;
            }
        }

        // Save changes
        wordDoc.MainDocumentPart.Document.Save();
    }
}

// Helper to add ordinal suffix to dates
string GetOrdinalSuffix(int day)
{
    if (day % 100 >= 11 && day % 100 <= 13) return "th"; // Handle 11th, 12th, 13th
    return (day % 10) switch
    {
        1 => "st",
        2 => "nd",
        3 => "rd",
        _ => "th"
    };
}

// Helper to convert numbers to words
string ConvertNumberToWords(int number)
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
