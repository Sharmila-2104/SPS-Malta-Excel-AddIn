using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace SPS_Malta_AaddIn
{
    public partial class MaltaRibbon
    {
        private void MaltaRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_Surname_Click(object sender, RibbonControlEventArgs e)
        {
            // Ask user for Excel columns
            string columnInput = Microsoft.VisualBasic.Interaction.InputBox(
                "Enter Excel column letters (e.g. E;O or J,E,R):",
                "Column Input", "E");

            if (string.IsNullOrEmpty(columnInput))
            {
                return;
            }

            // Split input columns by comma or semicolon
            var columns = columnInput.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);

            // Path to text file containing valid given names
            string textpath = @"C:\Swift_ProSys\Malta\SurName.txt";
            // Load valid names (case-insensitive)
            var givenNameList = File.ReadAllLines(textpath)
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .Select(line => line.Trim())
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            // Get active worksheet
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.Range usedRange = ws.UsedRange;
            int rowCount = usedRange.Rows.Count;

            foreach (string colLetter in columns)
            {
                int colIndex = ColumnLetterToNumber(colLetter);

                for (int row = 2; row <= rowCount; row++) // skip header row
                {
                    Excel.Range cell = (Excel.Range)ws.Cells[row, colIndex];
                    string value = cell.Value2?.ToString().Trim();

                    if (!string.IsNullOrEmpty(value))
                    {
                        // Reset font color
                        cell.Font.Color = ColorTranslator.ToOle(Color.Black);

                        // Split by space for multiple names
                        string[] nameGroups = value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        int pos = 1;

                        foreach (string group in nameGroups)
                        {
                            int length = group.Length;
                            bool allPartsValid = true;

                            // Split by hyphen for hyphenated names
                            string[] parts = group.Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);

                            foreach (string part in parts)
                            {
                                string trimmed = part.Trim();
                                if (!givenNameList.Contains(trimmed))
                                {
                                    allPartsValid = false;
                                    break;
                                }
                            }

                            // Highlight invalid names in red
                            if (!allPartsValid)
                            {
                                cell.Characters[pos, length].Font.Color = ColorTranslator.ToOle(Color.Red);
                            }

                            pos += length + 1; // Move cursor (+1 for space)
                        }
                    }
                }
            }
            MessageBox.Show("SurName validation completed!\nRed = Not in SurName list");
        }

        private void btn_Place_Click(object sender, RibbonControlEventArgs e)
        {
            // Ask user for Excel columns
            string columnInput = Microsoft.VisualBasic.Interaction.InputBox(
                "Enter Excel column letters (e.g. E;O or J,E,R):",
                "Column Input", "E");

            if (string.IsNullOrEmpty(columnInput))
            {
                return;
            }

            // Split input columns by comma or semicolon
            var columns = columnInput.Split(new[] { ',', ';' },
                StringSplitOptions.RemoveEmptyEntries);

            // Path to text file containing valid given names
            string textpath = @"C:\Swift_ProSys\Malta\Residence.txt";

            // Load valid names from text file into a HashSet (case-insensitive)
            var givenNameList = File.ReadAllLines(textpath)
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .Select(line => line.Trim())
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            // Get active worksheet
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.Range usedRange = ws.UsedRange;
            int rowCount = usedRange.Rows.Count;

            foreach (string colLetter in columns)
            {
                int colIndex = ColumnLetterToNumber(colLetter);

                for (int row = 2; row <= rowCount; row++) // skip header row
                {
                    Excel.Range cell = (Excel.Range)ws.Cells[row, colIndex];
                    string value = cell.Value2?.ToString().Trim();

                    if (!string.IsNullOrEmpty(value))
                    {
                        // Reset font color for entire cell
                        cell.Font.Color = ColorTranslator.ToOle(Color.Black);

                        // Split words by spaces
                        string[] words = value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                        int pos = 1; // Excel Characters are 1-based

                        foreach (string word in words)
                        {
                            int length = word.Length;
                            bool isValid = false;

                            // ✅ Case 1: word wrapped with hyphens (-Aabby-)
                            if (word.StartsWith("-") && word.EndsWith("-") && word.Length > 2)
                            {
                                string cleanWord = word.Trim('-'); // remove outer hyphens
                                if (givenNameList.Contains(cleanWord))
                                {
                                    isValid = true;
                                }
                            }
                            else
                            {
                                // ✅ Case 2: direct name match (without hyphen wrapping)
                                if (givenNameList.Contains(word))
                                {
                                    isValid = true;
                                }
                            }
                            // ❌ Highlight invalid names in red
                            if (!isValid)
                            {
                                cell.Characters[pos, length].Font.Color = ColorTranslator.ToOle(Color.Red);
                            }

                            // Move cursor position (+1 for space)
                            pos += length + 1;
                        }
                    }
                }
            }
            MessageBox.Show("Residence validation completed!\nRed = Not in Residence list");
        }

        private int ColumnLetterToNumber(string columnLetter)
        {
            int sum = 0;
            foreach (char c in columnLetter.ToUpper())
            {
                sum *= 26;
                sum += (c - 'A' + 1);
            }
            return sum;
        }

        private void btn_GivenName_Click(object sender, RibbonControlEventArgs e)
        {
            // Ask user for Excel columns
            string columnInput = Microsoft.VisualBasic.Interaction.InputBox(
                "Enter Excel column letters (e.g. E;O or J,E,R):",
                "Column Input", "E");

            if (string.IsNullOrEmpty(columnInput))
            {
                return;
            }

            // Split input columns by comma or semicolon
            var columns = columnInput.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);

            // Path to text file containing valid given names
            string textpath = @"C:\Swift_ProSys\Malta\GivenName.txt";

            // Load valid names (case-insensitive)
            var givenNameList = File.ReadAllLines(textpath)
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .Select(line => line.Trim())
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            // Get active worksheet
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.Range usedRange = ws.UsedRange;
            int rowCount = usedRange.Rows.Count;

            foreach (string colLetter in columns)
            {
                int colIndex = ColumnLetterToNumber(colLetter);

                for (int row = 2; row <= rowCount; row++) // skip header row
                {
                    Excel.Range cell = (Excel.Range)ws.Cells[row, colIndex];
                    string value = cell.Value2?.ToString().Trim();

                    if (!string.IsNullOrEmpty(value))
                    {
                        // Reset font color
                        cell.Font.Color = ColorTranslator.ToOle(Color.Black);

                        // Split by space for multiple names
                        string[] nameGroups = value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        int pos = 1;

                        foreach (string group in nameGroups)
                        {
                            int length = group.Length;
                            bool allPartsValid = true;

                            // Split by hyphen for hyphenated names
                            string[] parts = group.Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);

                            foreach (string part in parts)
                            {
                                string trimmed = part.Trim();
                                if (!givenNameList.Contains(trimmed))
                                {
                                    allPartsValid = false;
                                    break;
                                }
                            }

                            // Highlight invalid names in red
                            if (!allPartsValid)
                            {
                                cell.Characters[pos, length].Font.Color = ColorTranslator.ToOle(Color.Red);
                            }

                            pos += length + 1; // Move cursor (+1 for space)
                        }
                    }
                }
            }
            MessageBox.Show("✅ GivenName validation completed!\nRed = Not in GivenName list");
        }

        private void btn_FileNameImageNumber_Click(object sender, RibbonControlEventArgs e)
        {
            // Ask for Excel column letters
            string columnInput = Microsoft.VisualBasic.Interaction.InputBox(
                "Enter Excel column letter for Folder ID (e.g. B):",
                "Column Input", "B");

            string columnInput2 = Microsoft.VisualBasic.Interaction.InputBox(
                "Enter Excel column letter for Image ID (e.g. C):",
                "Column Input", "C");

            if (string.IsNullOrWhiteSpace(columnInput) || string.IsNullOrWhiteSpace(columnInput2))
                return;

            // Convert letters to numbers
            int folderCol = ColumnLetterToNumber(columnInput);
            int imageCol = ColumnLetterToNumber(columnInput2);

            // Path to reference file
            string textpath = @"C:\Swift_ProSys\Malta\Folder_Image ID.txt";

            if (!File.Exists(textpath))
            {
                MessageBox.Show("Database file not found:\n" + textpath);
                return;
            }

            // Read database lines and store valid combinations in HashSet
            var validPairs = File.ReadAllLines(textpath)
                .Where(line => !string.IsNullOrWhiteSpace(line) && line.Contains("|"))
                .Select(line => line.Trim())
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            // Get active Excel sheet
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.Range usedRange = ws.UsedRange;
            int rowCount = usedRange.Rows.Count;

            // Start checking row by row
            for (int row = 2; row <= rowCount; row++) // skip header
            {
                Excel.Range folderCell = (Excel.Range)ws.Cells[row, folderCol];
                Excel.Range imageCell = (Excel.Range)ws.Cells[row, imageCol];

                string folderValue = folderCell.Value2?.ToString().Trim() ?? "";
                string imageValue = imageCell.Value2?.ToString().Trim() ?? "";

                // Reset colors first
                folderCell.Font.Color = ColorTranslator.ToOle(Color.Black);
                imageCell.Font.Color = ColorTranslator.ToOle(Color.Black);

                if (!string.IsNullOrEmpty(folderValue) && !string.IsNullOrEmpty(imageValue))
                {
                    string combo = folderValue + "|" + imageValue;

                    // If not found in database file → mark red
                    if (!validPairs.Contains(combo))
                    {
                        folderCell.Font.Color = ColorTranslator.ToOle(Color.Red);
                        imageCell.Font.Color = ColorTranslator.ToOle(Color.Red);
                    }
                }
            }
            MessageBox.Show("Folder/Image validation completed!\nRed = Not found in database");
        }

        private void btn_FourDigit_Click(object sender, RibbonControlEventArgs e)
        {
            //// Ask user for Excel columns
            //string columnInput = Microsoft.VisualBasic.Interaction.InputBox(
            //    "Enter Excel column letters (e.g. E;O or J,E,R):",
            //    "Column Input", "E");

            //if (string.IsNullOrEmpty(columnInput))
            //{
            //    return;
            //}

            //// Split input columns by comma or semicolon
            //var columns = columnInput.Split(new[] { ',', ';' },
            //    StringSplitOptions.RemoveEmptyEntries);

            //// Get active worksheet
            //Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            //Excel.Range usedRange = ws.UsedRange;
            //int rowCount = usedRange.Rows.Count;

            //foreach (string colLetter in columns)
            //{
            //    int colIndex = ColumnLetterToNumber(colLetter);

            //    for (int row = 2; row <= rowCount; row++) // skip header row
            //    {
            //        Excel.Range cell = (Excel.Range)ws.Cells[row, colIndex];
            //        string value = cell.Value2?.ToString().Trim();

            //        if (!string.IsNullOrEmpty(value))
            //        {
            //            // Reset to black
            //            cell.Font.Color = ColorTranslator.ToOle(Color.Black);

            //            // ✅ Must be 4-digit and within 1520–1990
            //            if (!System.Text.RegularExpressions.Regex.IsMatch(value, @"^\d{4}$"))
            //            {
            //                cell.Font.Color = ColorTranslator.ToOle(Color.Red);
            //            }
            //            else
            //            {
            //                int num = int.Parse(value);
            //                if (num < 1520 || num > 1990)
            //                {
            //                    cell.Font.Color = ColorTranslator.ToOle(Color.Red);
            //                }
            //            }
            //        }
            //    }
            //}
            //MessageBox.Show("Four-digit and range validation completed!\nRed = Invalid (not 4 digits or out of 1520–1990)");

            // Ask user for Excel column letters (e.g., "E" or "E;O")
            string columnInput = Microsoft.VisualBasic.Interaction.InputBox(
                "Enter Excel column letters (e.g. E;O or J,E,R):",
                "Column Input", "E");

            if (string.IsNullOrEmpty(columnInput))
                return;

            // Split input columns by comma or semicolon
            var columns = columnInput.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);

            // Get active worksheet
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.Range usedRange = ws.UsedRange;
            int lastRow = usedRange.Rows.Count;

            foreach (var col in columns)
            {
                string column = col.Trim().ToUpper();

                for (int row = 1; row <= lastRow; row++)
                {
                    Excel.Range cell = ws.Cells[row, column] as Excel.Range;

                    if (cell != null && cell.Value2 != null)
                    {
                        string cellValue = cell.Value2.ToString().Trim();

                        // Check total length (including hyphens or slashes)
                        int length = cellValue.Length;

                        if (length > 10)
                        {
                            // Mark red if more than 10 characters
                            cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                        else
                        {
                            // Clear any previous color
                            cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                        }
                    }
                }
            }

            MessageBox.Show("Cells with more than 10 characters are marked in red.",
                "Process Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btn_Gender_Click(object sender, RibbonControlEventArgs e)
        {
            // Ask user for Excel columns
            string columnInput = Microsoft.VisualBasic.Interaction.InputBox(
                "Enter Excel column letters (e.g. E;O or J,E,R):",
                "Column Input", "E");

            if (string.IsNullOrEmpty(columnInput))
            {
                return;
            }

            // Split input columns by comma or semicolon
            var columns = columnInput.Split(new[] { ',', ';' },
                StringSplitOptions.RemoveEmptyEntries);

            // Get active worksheet
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.Range usedRange = ws.UsedRange;
            int rowCount = usedRange.Rows.Count;

            foreach (string colLetter in columns)
            {
                int colIndex = ColumnLetterToNumber(colLetter.Trim());

                for (int row = 2; row <= rowCount; row++) // skip header
                {
                    Excel.Range cell = (Excel.Range)ws.Cells[row, colIndex];
                    string value = cell.Value2?.ToString().Trim().ToUpper();

                    if (!string.IsNullOrEmpty(value))
                    {
                        // Reset color to black
                        cell.Font.Color = ColorTranslator.ToOle(Color.Black);

                        // ✅ Check gender validity: must be M or F
                        if (value != "M" && value != "F")
                        {
                            // ❌ Invalid → highlight red
                            cell.Font.Color = ColorTranslator.ToOle(Color.Red);
                        }
                    }
                }
            }

            MessageBox.Show("Gender validation completed!\nRed = Invalid (not 'M' or 'F')");

        }

        private void btn_MaritalStatus_Click(object sender, RibbonControlEventArgs e)
        {
            // Ask user for Excel columns
            string columnInput = Microsoft.VisualBasic.Interaction.InputBox(
                "Enter Excel column letters (e.g. E;O or J,E,R):",
                "Column Input", "E");

            if (string.IsNullOrEmpty(columnInput))
            {
                return;
            }

            // Split input columns by comma or semicolon
            var columns = columnInput.Split(new[] { ',', ';' },
                StringSplitOptions.RemoveEmptyEntries);

            // Get active worksheet
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.Range usedRange = ws.UsedRange;
            int rowCount = usedRange.Rows.Count;

            foreach (string colLetter in columns)
            {
                int colIndex = ColumnLetterToNumber(colLetter.Trim());

                for (int row = 2; row <= rowCount; row++) // skip header row
                {
                    Excel.Range cell = (Excel.Range)ws.Cells[row, colIndex];
                    string value = cell.Value2?.ToString().Trim().ToUpper();

                    if (!string.IsNullOrEmpty(value))
                    {
                        // Reset color to black first
                        cell.Font.Color = ColorTranslator.ToOle(Color.Black);

                        // ✅ Check marital status validity: must be S or W
                        if (value != "S" && value != "W")
                        {
                            // ❌ Invalid → highlight red
                            cell.Font.Color = ColorTranslator.ToOle(Color.Red);
                        }
                    }
                }
            }

            MessageBox.Show("Marital Status validation completed!\nRed = Invalid (not 'S' or 'W')");

        }

        private void btn_ExcelMerge_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;

                using (FolderBrowserDialog fbd = new FolderBrowserDialog())
                {
                    fbd.Description = "Select a folder that contains Excel files";
                    fbd.RootFolder = Environment.SpecialFolder.MyComputer;

                    if (fbd.ShowDialog() == DialogResult.OK)
                    {
                        string folderPath = fbd.SelectedPath;
                        string[] excelFiles = Directory.GetFiles(folderPath, "*.xls*");

                        if (excelFiles.Length == 0)
                        {
                            MessageBox.Show("No Excel files found in the selected folder.");
                            return;
                        }

                        // Create new workbook to merge data
                        Excel.Workbook mergeWb = app.Workbooks.Add();
                        Excel.Worksheet mergeWs = (Excel.Worksheet)mergeWb.Sheets[1];

                        int destRow = 1;
                        bool headerCopied = false;

                        foreach (string file in excelFiles)
                        {
                            Excel.Workbook wb = app.Workbooks.Open(file);
                            Excel.Worksheet ws = wb.Sheets[1]; // Use first sheet

                            Excel.Range usedRange = ws.UsedRange;
                            int rows = usedRange.Rows.Count;
                            int cols = usedRange.Columns.Count;

                            // Determine source range
                            int startRow = headerCopied ? 2 : 1; // Skip header if already copied
                            int rowsToCopy = rows - (headerCopied ? 1 : 0);

                            if (rowsToCopy > 0)
                            {
                                Excel.Range sourceRange = ws.Range[ws.Cells[startRow, 1], ws.Cells[startRow + rowsToCopy - 1, cols]];
                                Excel.Range destRange = mergeWs.Range[mergeWs.Cells[destRow, 1], mergeWs.Cells[destRow + rowsToCopy - 1, cols]];
                                sourceRange.Copy(destRange);
                                destRow += rowsToCopy;
                            }

                            headerCopied = true;
                            wb.Close(false);
                        }

                        mergeWs.Columns.AutoFit();
                        MessageBox.Show("✅ All Excel files merged successfully!\nHeaders included only once.\nSource folder: " + folderPath);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void btn_Split_Click(object sender, RibbonControlEventArgs e)
        {

            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Workbook activeWb = app.ActiveWorkbook;
            Excel.Worksheet ws = activeWb.ActiveSheet;

            Excel.Range usedRange = ws.UsedRange;
            int totalRows = usedRange.Rows.Count;

            // Ask for column (e.g. "E")
            string columnInput = Microsoft.VisualBasic.Interaction.InputBox(
                "Enter Excel column letter (e.g., E)", "Column Input", "A");

            if (string.IsNullOrWhiteSpace(columnInput))
            {
                MessageBox.Show("No column provided.");
                return;
            }

            int colIndex = ws.Range[columnInput + "1"].Column;

            // Choose folder to save files
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "Select folder to save split files (each in its own folder)";
                fbd.RootFolder = Environment.SpecialFolder.MyComputer;

                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    string baseFolder = fbd.SelectedPath;

                    // Collect unique values
                    HashSet<string> uniqueValues = new HashSet<string>();
                    for (int row = 2; row <= totalRows; row++) // skip header row
                    {
                        var cellValue = (ws.Cells[row, colIndex] as Excel.Range).Value2;
                        if (cellValue != null)
                        {
                            uniqueValues.Add(cellValue.ToString());
                        }
                    }

                    // Process each unique value
                    foreach (string value in uniqueValues)
                    {
                        Excel.Workbook newWb = app.Workbooks.Add();
                        Excel.Worksheet newWs = newWb.ActiveSheet;

                        int pasteRow = 1;

                        // Copy header
                        Excel.Range headerRow = ws.Rows[1];
                        headerRow.Copy(newWs.Rows[pasteRow]);
                        pasteRow++;

                        // Copy matching rows
                        for (int row = 2; row <= totalRows; row++)
                        {
                            var cellValue = (ws.Cells[row, colIndex] as Excel.Range).Value2;
                            if (cellValue != null && cellValue.ToString() == value)
                            {
                                Excel.Range sourceRow = ws.Rows[row];
                                Excel.Range destRow = newWs.Rows[pasteRow];
                                sourceRow.Copy(destRow);
                                pasteRow++;
                            }
                        }

                        // AutoFit all columns
                        newWs.Columns.AutoFit();

                        // Clean up value for folder/file name
                        string safeValue = string.Join("_", value.Split(Path.GetInvalidFileNameChars()));

                        //// Create subfolder for this value
                        //string subFolder = Path.Combine(baseFolder, safeValue);
                        //Directory.CreateDirectory(subFolder);

                        //// Save inside its subfolder
                        //string savePath = Path.Combine(subFolder, $"{safeValue}.xlsx");
                        //newWb.SaveAs(savePath);
                        //newWb.Close(false);

                        // Save directly inside the chosen folder
                        string savePath = Path.Combine(baseFolder, $"{safeValue}.xlsx");
                        newWb.SaveAs(savePath);
                        newWb.Close(false);
                    }

                    MessageBox.Show("✅ Split completed!\nFiles saved in individual folders under:\n" + baseFolder);
                }
            }
        }

        private void btn_BirthToBaptisms_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Worksheet ws = app.ActiveSheet;

                // Ask user for 2 column letters
                string colInput1 = Microsoft.VisualBasic.Interaction.InputBox(
                    "Enter Birth Date column letter (e.g., A):", "Column Input", "A");
                string colInput2 = Microsoft.VisualBasic.Interaction.InputBox(
                    "Enter Baptism Date column letter (e.g., B):", "Column Input", "B");

                if (string.IsNullOrWhiteSpace(colInput1) || string.IsNullOrWhiteSpace(colInput2)) return;

                Excel.Range usedRange = ws.UsedRange;
                int lastRow = usedRange.Rows.Count;

                for (int row = 1; row <= lastRow; row++)
                {
                    Excel.Range deathCell = ws.Range[colInput1 + row];
                    Excel.Range burialCell = ws.Range[colInput2 + row];

                    if (deathCell.Value2 != null && burialCell.Value2 != null)
                    {
                        DateTime deathDate = DateTime.MinValue;
                        DateTime burialDate = DateTime.MinValue;

                        if (DateTime.TryParse(deathCell.Value2.ToString(), out deathDate) &&
                            DateTime.TryParse(burialCell.Value2.ToString(), out burialDate))
                        {
                            if (burialDate < deathDate)
                            {
                                // Burial before Death → Highlight Red
                                burialCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                deathCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                                //burialCell.Font.Color = ColorTranslator.ToOle(Color.Red);
                                //deathCell.Font.Color = ColorTranslator.ToOle(Color.Red);
                            }
                            else
                            {
                                // Valid → reset color
                                burialCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                                deathCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                                //burialCell.Font.Color = ColorTranslator.ToOle(Color.White);
                                //deathCell.Font.Color = ColorTranslator.ToOle(Color.White);
                            }
                        }
                    }
                }
                MessageBox.Show("Death vs Burial validation complete!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void btn_DeathToBurial_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Worksheet ws = app.ActiveSheet;

                // Ask user for 2 column letters
                string colInput1 = Microsoft.VisualBasic.Interaction.InputBox(
                    "Enter Death Date column letter (e.g., A):", "Column Input", "A");
                string colInput2 = Microsoft.VisualBasic.Interaction.InputBox(
                    "Enter Burial Date column letter (e.g., B):", "Column Input", "B");

                if (string.IsNullOrWhiteSpace(colInput1) || string.IsNullOrWhiteSpace(colInput2)) return;

                Excel.Range usedRange = ws.UsedRange;
                int lastRow = usedRange.Rows.Count;

                for (int row = 1; row <= lastRow; row++)
                {
                    Excel.Range deathCell = ws.Range[colInput1 + row];
                    Excel.Range burialCell = ws.Range[colInput2 + row];

                    if (deathCell.Value2 != null && burialCell.Value2 != null)
                    {
                        DateTime deathDate = DateTime.MinValue;
                        DateTime burialDate = DateTime.MinValue;

                        if (DateTime.TryParse(deathCell.Value2.ToString(), out deathDate) &&
                            DateTime.TryParse(burialCell.Value2.ToString(), out burialDate))
                        {
                            if (burialDate < deathDate)
                            {
                                // Burial before Death → Highlight Red
                                burialCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                deathCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                //burialCell.Font.Color = ColorTranslator.ToOle(Color.Red);
                                //deathCell.Font.Color = ColorTranslator.ToOle(Color.Red);
                            }
                            else
                            {
                                // Valid → reset color
                                burialCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                                deathCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                                //burialCell.Font.Color = ColorTranslator.ToOle(Color.White);
                                //deathCell.Font.Color = ColorTranslator.ToOle(Color.White);
                            }
                        }
                    }
                }
                MessageBox.Show("Death vs Burial validation complete!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void btn_MarriageToBanns_Click(object sender, RibbonControlEventArgs e)
        {
            //try
            //{
            //    Excel.Application app = Globals.ThisAddIn.Application;
            //    Excel.Worksheet ws = app.ActiveSheet;

            //    // Ask user for column letters
            //    string colInput1 = Microsoft.VisualBasic.Interaction.InputBox(
            //        "Enter Marriage Date column letter (e.g., A):", "Column Input", "A");
            //    string colInput2 = Microsoft.VisualBasic.Interaction.InputBox(
            //        "Enter Banns Date column letter (e.g., B):", "Column Input", "B");

            //    Excel.Range usedRange = ws.UsedRange;
            //    int lastRow = usedRange.Rows.Count;

            //    for (int i = 2; i <= lastRow; i++) // assuming row 1 has headers
            //    {
            //        string marriageCol = colInput1 + i;
            //        string bannsCol = colInput2 + i;

            //        Excel.Range marriageCell = ws.Range[marriageCol];
            //        Excel.Range bannsCell = ws.Range[bannsCol];

            //        DateTime marriageDate, bannsDate;

            //        // Try parsing both dates safely
            //        DateTime.TryParse(Convert.ToString(marriageCell.Value), out marriageDate);
            //        DateTime.TryParse(Convert.ToString(bannsCell.Value), out bannsDate);

            //        // Compare dates only if both are valid
            //        if (marriageDate != DateTime.MinValue && bannsDate != DateTime.MinValue)
            //        {
            //            if (marriageDate < bannsDate)
            //            {
            //                // Marriage Date is earlier than Banns Date → highlight it or show message
            //                marriageCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            //                bannsCell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            //            }
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    System.Windows.Forms.MessageBox.Show("Error: " + ex.Message);
            //}

            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Worksheet ws = app.ActiveSheet;

            string marriageCol = Microsoft.VisualBasic.Interaction.InputBox(
                "Enter Marriage Date column letter (e.g., A):", "Column Input", "A");
            string bans1Col = Microsoft.VisualBasic.Interaction.InputBox(
                "Enter Banns 1 Date column letter (e.g., B):", "Column Input", "B");
            string bans2Col = Microsoft.VisualBasic.Interaction.InputBox(
                "Enter Banns 2 Date column letter (e.g., C):", "Column Input", "C");
            string bans3Col = Microsoft.VisualBasic.Interaction.InputBox(
                "Enter Banns 3 Date column letter (e.g., D):", "Column Input", "D");

            Excel.Range usedRange = ws.UsedRange;
            int totalRows = usedRange.Rows.Count;

            int marriageIndex = ExcelColToIndex(marriageCol);
            int bans1Index = ExcelColToIndex(bans1Col);
            int bans2Index = ExcelColToIndex(bans2Col);
            int bans3Index = ExcelColToIndex(bans3Col);

            for (int i = 2; i <= totalRows; i++) // row 1 has headers
            {
                string marriageStr = ws.Cells[i, marriageIndex].Text.ToString();
                string b1Str = ws.Cells[i, bans1Index].Text.ToString();
                string b2Str = ws.Cells[i, bans2Index].Text.ToString();
                string b3Str = ws.Cells[i, bans3Index].Text.ToString();

                DateTime marriageDate, bans1, bans2, bans3;
                bool hasMarriage = DateTime.TryParse(marriageStr, out marriageDate);
                bool hasB1 = DateTime.TryParse(b1Str, out bans1);
                bool hasB2 = DateTime.TryParse(b2Str, out bans2);
                bool hasB3 = DateTime.TryParse(b3Str, out bans3);

                // Reset background before checking (optional, if you rerun multiple times)
                ws.Cells[i, marriageIndex].Interior.ColorIndex = 0;
                ws.Cells[i, bans1Index].Interior.ColorIndex = 0;
                ws.Cells[i, bans2Index].Interior.ColorIndex = 0;
                ws.Cells[i, bans3Index].Interior.ColorIndex = 0;

                // Highlight invalid date formats only
                if (!hasMarriage && !string.IsNullOrWhiteSpace(marriageStr))
                    ws.Cells[i, marriageIndex].Interior.Color = System.Drawing.Color.Red;
                if (!hasB1 && !string.IsNullOrWhiteSpace(b1Str))
                    ws.Cells[i, bans1Index].Interior.Color = System.Drawing.Color.Red;
                if (!hasB2 && !string.IsNullOrWhiteSpace(b2Str))
                    ws.Cells[i, bans2Index].Interior.Color = System.Drawing.Color.Red;
                if (!hasB3 && !string.IsNullOrWhiteSpace(b3Str))
                    ws.Cells[i, bans3Index].Interior.Color = System.Drawing.Color.Red;

                // Order validation (check only if both dates exist)
                if (hasB1 && hasB2 && bans1 > bans2)
                {
                    ws.Cells[i, bans1Index].Interior.Color = System.Drawing.Color.Red;
                    ws.Cells[i, bans2Index].Interior.Color = System.Drawing.Color.Red;
                }
                if (hasB2 && hasB3 && bans2 > bans3)
                {
                    ws.Cells[i, bans2Index].Interior.Color = System.Drawing.Color.Red;
                    ws.Cells[i, bans3Index].Interior.Color = System.Drawing.Color.Red;
                }
                if (hasB3 && hasMarriage && bans3 > marriageDate)
                {
                    ws.Cells[i, bans3Index].Interior.Color = System.Drawing.Color.Red;
                    ws.Cells[i, marriageIndex].Interior.Color = System.Drawing.Color.Red;
                }
            }

            MessageBox.Show("Validation completed!");
        }
        private int ExcelColToIndex(string col)
        {
            col = col.ToUpper();
            int sum = 0;
            for (int i = 0; i < col.Length; i++)
            {
                sum *= 26;
                sum += (col[i] - 'A' + 1);
            }
            return sum;
        }
        private void btn_Headercheck_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // Ask user which template to check
                string selected = Microsoft.VisualBasic.Interaction.InputBox(
                    "Enter Template Type:\n(Birth_Baptism, Confirmations, Death_Burials, Marriage, Status_Animarum)",
                    "Select Template", "Birth_Baptism");

                if (string.IsNullOrEmpty(selected))
                    return;

                selected = selected.Trim().Replace(" ", "_");

                // ✅ Define expected headers for each template
                Dictionary<string, string[]> templateHeaders = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
                {
                    {
                        "Birth_Baptism", new string[]
                        {
                            "Folder Name", "Image Number", "Page or folio", "Entry Number",
                            "Birth Date", "Baptism Date", "Forenames", "Surname", "Sex",
                            "Father's Forenames", "Father's Surname",
                            "Mother's Forenames", "Mother's Maiden Surname",
                            "Godparent 1 Forenames", "Godparent 1 Surname",
                            "Godparent 2 Forenames", "Godparent 2 Surname",
                            "Godparent 3 Forenames", "Godparent 3 Surname"
                        }
                    },
                    {
                        "Confirmations", new string[]
                        {
                            "Folder Name", "Image Number", "Page or folio", "Entry Number",
                            "Confirmation Date", "Forenames", "Surname", "Sex", "Age",
                            "Father's Forenames", "Father's Surname", "Mother's Forenames",
                            "Mother's Maiden Surname", "Godparent Forenames", "Godparent Surname",
                            "Godparent Place"
                        }
                    },
                    {
                        "Death_Burials", new string[]
                        {
                            "Folder Name", "Image Number", "Page or folio", "Entry Number",
                            "Forenames", "Surname", "Sex", "Age", "Father's Forenames", "Father's Surname",
                            "Mother's Forenames","Mother's Surname","Marital Status","Spouse's Forenames",
                            "Spouse's Surname","Death Date","Burial Date","Witness 1 Forenames","Witness 1 Surname",
                            "Witness 2 Forenames","Witness 2 Surname","Witness 3 Forenames","Witness 3 Surname"
                        }
                    },
                    {
                        "Marriage", new string[]
                        {
                            "Folder Name", "Image Number", "Page or folio", "Entry Number",
                            "Groom's Forenames", "Groom's Surname", "Groom's Age","Groom's Marital Status",
                            "Groom's Father's Forenames","Groom's Father's Surname","Groom's Mother's Forenames",
                            "Groom's Mother's Surname","Marriage Date","Banns 1 Date","Banns 2 Date","Banns 3 Date",
                            "Bride's Forename","Bride's Surname","Bride's Age","Bride's Marital Status","Bride's Father's Forenames",
                            "Bride's Father's Surname","Bride's Mother's Forenames","Bride's Mother's Surname","Witness 1 Forenames",
                            "Witness 1 Surname","Witness 2 Forenames","Witness 2 Surname","Witness 3 Forenames","Witness 3 Surname"

                        }
                    },
                    {
                        "Status_Animarum", new string[]
                        {
                            "Folder Name", "Image Number", "Page or folio", "Entry Number","Year",
                            "Forenames", "Surname", "Maiden surname","Sex","Age","Father's Forenames",
                            "Father's Surname","Mother's Forenames","Mother's Maiden Surname","Address"
                        }
                    }
                };

                if (!templateHeaders.ContainsKey(selected))
                {
                    MessageBox.Show("Invalid template type entered!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string[] expectedHeaders = templateHeaders[selected];

                // Get active worksheet and headers from first row
                Excel.Application app = Globals.ThisAddIn.Application;
                Excel.Worksheet ws = app.ActiveSheet as Excel.Worksheet;
                Excel.Range usedRange = ws.UsedRange;
                int colCount = usedRange.Columns.Count;

                List<string> currentHeaders = new List<string>();
                for (int c = 1; c <= colCount; c++)
                {
                    Excel.Range cell = (Excel.Range)ws.Cells[1, c];
                    string value = cell.Value2?.ToString()?.Trim() ?? "";
                    currentHeaders.Add(value);
                }

                // Compare headers
                int maxCols = Math.Max(expectedHeaders.Length, currentHeaders.Count);
                for (int i = 0; i < maxCols; i++)
                {
                    Excel.Range cell = (Excel.Range)ws.Cells[1, i + 1];
                    cell.Font.Color = ColorTranslator.ToOle(Color.Black); // reset color

                    string expected = i < expectedHeaders.Length ? expectedHeaders[i] : "";
                    string current = i < currentHeaders.Count ? currentHeaders[i] : "";

                    if (!string.Equals(expected, current, StringComparison.OrdinalIgnoreCase))
                    {
                        // ❌ Highlight mismatched header
                        cell.Font.Color = ColorTranslator.ToOle(Color.Red);
                    }
                }

                MessageBox.Show($"{selected} header validation completed!\nRed = Mismatched or missing header.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }


            //try
            //{
            //    // ✅ Expected headers for Birth Baptism template
            //    string[] expectedHeaders = new string[]
            //    {
            //    "Folder Name", "Image Number", "Page or folio", "Entry Number",
            //    "Birth Date", "Baptism Date", "Forenames", "Surname", "Sex",
            //    "Father's Forenames", "Father's Surname",
            //    "Mother's Forenames", "Mother's Maiden Surname",
            //    "Godparent 1 Forenames", "Godparent 1 Surname",
            //    "Godparent 2 Forenames", "Godparent 2 Surname",
            //    "Godparent 3 Forenames", "Godparent 3 Surname"
            //    };

            //    Excel.Application app = Globals.ThisAddIn.Application;
            //    Excel.Worksheet ws = app.ActiveSheet as Excel.Worksheet;
            //    Excel.Range usedRange = ws.UsedRange;
            //    int colCount = usedRange.Columns.Count;

            //    // Read headers from the first row of the active sheet
            //    List<string> currentHeaders = new List<string>();
            //    for (int c = 1; c <= colCount; c++)
            //    {
            //        Excel.Range cell = (Excel.Range)ws.Cells[1, c];
            //        string value = cell.Value2?.ToString()?.Trim() ?? "";
            //        currentHeaders.Add(value);
            //    }

            //    int maxCols = Math.Max(expectedHeaders.Length, currentHeaders.Count);

            //    for (int i = 0; i < maxCols; i++)
            //    {
            //        Excel.Range cell = (Excel.Range)ws.Cells[1, i + 1];
            //        cell.Font.Color = ColorTranslator.ToOle(Color.Black); // reset color

            //        string expected = i < expectedHeaders.Length ? expectedHeaders[i] : "";
            //        string current = i < currentHeaders.Count ? currentHeaders[i] : "";

            //        if (!string.Equals(expected, current, StringComparison.OrdinalIgnoreCase))
            //        {
            //            // ❌ Highlight mismatched header in red
            //            cell.Font.Color = ColorTranslator.ToOle(Color.Red);
            //        }
            //    }

            //    MessageBox.Show("Birth Baptism header validation completed!\nRed = Mismatched or missing header.");
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Error: " + ex.Message);
            //}
        }

        private void btn_Trim_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
                Excel.Range usedRange = ws.UsedRange;
                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;

                for (int row = 1; row <= rowCount; row++)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        Excel.Range cell = (Excel.Range)(ws.Cells[row, col]);

                        if (cell.Value2 != null)
                        {
                            string value = cell.Text;  //use .Text to preserve formatting exactly
                            value = value.Trim();

                            while (value.Contains("  "))
                                value = value.Replace("  ", " ");

                            //  Force Text format before writing
                            cell.NumberFormat = "@";
                            cell.Value2 = value;
                        }

                        // Remove Wrap Text
                        cell.WrapText = false;

                        //  Remove Merge & Center
                        if (cell.MergeCells)
                        {
                            cell.MergeArea.UnMerge();
                        }

                        // Reset alignment
                        cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignGeneral;
                    }
                }

                MessageBox.Show("Trim done ✅ (leading zeros preserved, wrap/merge cleared)");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            //try
            //{
            //    Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            //    Excel.Range usedRange = ws.UsedRange;
            //    int rowCount = usedRange.Rows.Count;
            //    int colCount = usedRange.Columns.Count;

            //    for (int row = 1; row <= rowCount; row++)
            //    {
            //        for (int col = 1; col <= colCount; col++)
            //        {
            //            Excel.Range cell = (Excel.Range)(ws.Cells[row, col]);

            //            if (cell.Value2 != null)
            //            {
            //                string value = cell.Value2.ToString();

            //                // Trim spaces
            //                value = value.Trim();
            //                while (value.Contains("  "))
            //                    value = value.Replace("  ", " ");

            //                cell.Value2 = value;
            //            }

            //            // Remove Wrap Text
            //            cell.WrapText = false;

            //            // 🔹 Remove Merge & Center
            //            if (cell.MergeCells)
            //            {
            //                cell.MergeArea.UnMerge(); // Unmerge all merged cells
            //            }
            //            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignGeneral; // Reset alignment
            //        }
            //    }

            //    MessageBox.Show("Trim, Wrap Text removed, Merge & Center cleared!");
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }
    }
}
