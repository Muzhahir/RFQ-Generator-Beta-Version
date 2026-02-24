using ClosedXML.Excel;
using RFQ_Generator_System.Repositories;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RFQ_Generator_System.Services
{
    public class ExcelGenerationService
    {
        private readonly ClientRepo clientRepo;
        private readonly RFQItemRepo rfqItemRepo;

        private static class Placeholders
        {
            public const string ClientName = "client_name";
            public const string Date = "_date";
            public const string RFQCode = "rfq_code";
            public const string QuoteCode = "quote_code";
            public const string Validity = "_validity";
            public const string PaymentTerm = "payment_term";
            public const string DeliveryTerm = "delivery_term";
            public const string DeliveryPoint = "delivery_point";
            public const string ItemNo = "_num";
            public const string ItemDesc = "_desc";
            public const string DeliveryTime = "delivery_time";
            public const string Quantity = "_qty";
            public const string UnitPrice = "unit_price";
            public const string TotalPrice = "total_price";
            public const string Subtotal = "sub_total";
            public const string Discount = "_discount";
            public const string SummaryTotal = "summary_total_price";
            public const string HeaderDeliveryTime = "header_delivery_time";
        }

        public ExcelGenerationService()
        {
            clientRepo = new ClientRepo();
            rfqItemRepo = new RFQItemRepo();
        }

        /// <summary>
        /// Generate priced RFQ Excel (default behavior)
        /// </summary>
        public string GenerateRFQExcel(string templatePath, string outputPath, RFQ rfq, List<RFQItem> items, int templateId)
        {
            return GenerateRFQExcel(templatePath, outputPath, rfq, items, templateId, isPriced: true);
        }

        /// <summary>
        /// Generate RFQ Excel with option for priced or unpriced version
        /// RETURNS the version suffix used (e.g., "PRICED", "UNPRICED", "COMMERCIAL", "TECHNICAL")
        /// </summary>
        /// <param name="isPriced">True for priced version, False for unpriced/technical version</param>
        public string GenerateRFQExcel(string templatePath, string outputPath, RFQ rfq, List<RFQItem> items, int templateId, bool isPriced)
        {
            string fullTemplatePath = GetFullTemplatePath(templatePath);

            using (var workbook = new XLWorkbook(fullTemplatePath))
            {
                var worksheet = workbook.Worksheet(1);

                // DETECT TEMPLATE TERMINOLOGY FIRST
                var (pricedTerm, unpricedTerm) = DetectTemplateTerminology(worksheet);

                // Determine which term to use based on isPriced
                string versionSuffix = isPriced ? pricedTerm : unpricedTerm;

                var client = clientRepo.GetClientById(rfq.ClientId);
                string clientName = client?.ClientName ?? "";

                var itemStartCell = FindPlaceholderCell(worksheet, Placeholders.ItemNo);

                // Determine effective currency BEFORE any operations
                string effectiveCurrency = rfq.Currency ?? "RM";

                // DEBUG: Show what terminology and currency we're using
                System.Diagnostics.Debug.WriteLine($"[ExcelService] Template uses: {pricedTerm}/{unpricedTerm}");
                System.Diagnostics.Debug.WriteLine($"[ExcelService] Version suffix: {versionSuffix}");
                System.Diagnostics.Debug.WriteLine($"[ExcelService] Received RFQ.Currency: '{rfq.Currency ?? "NULL"}'");
                System.Diagnostics.Debug.WriteLine($"[ExcelService] Effective Currency: '{effectiveCurrency}'");
                System.Diagnostics.Debug.WriteLine($"[ExcelService] Is Priced: {isPriced}");

                // Convert PRICED/COMMERCIAL to UNPRICED/TECHNICAL if generating unpriced version
                if (!isPriced)
                {
                    ConvertToUnpricedVersion(worksheet, pricedTerm, unpricedTerm);
                }

                // Fill header fields including header_delivery_time placeholder
                FillHeaderFields(worksheet, rfq, clientName, items);

                if (itemStartCell != null)
                {
                    FillItems(worksheet, items, itemStartCell, effectiveCurrency, isPriced);
                }

                // For unpriced version, discount is always 0
                decimal effectiveDiscount = isPriced ? rfq.Discount : 0;
                FillSummary(worksheet, items, effectiveDiscount, isPriced);

                // Final currency replacement to catch any remaining instances
                ReplaceCurrencyInWorksheet(worksheet, effectiveCurrency);

                // Apply bottom borders to item rows at page breaks (except for OGIT and MA templates)
                ApplyBottomBordersAtPageBreaks(worksheet, fullTemplatePath);

                workbook.SaveAs(outputPath);

                // RETURN the version suffix that was used
                return versionSuffix;
            }
        }

        /// <summary>
        /// Generate BOTH priced and unpriced Excel files at the same time
        /// Returns tuple of (pricedPath, unpricedPath, pricedSuffix, unpricedSuffix)
        /// </summary>
        public (string pricedPath, string unpricedPath, string pricedSuffix, string unpricedSuffix) GenerateBothVersionsExcel(
            string templatePath,
            string baseOutputPath,
            RFQ rfq,
            List<RFQItem> items,
            int templateId)
        {
            // Generate paths for both versions
            string directory = Path.GetDirectoryName(baseOutputPath);
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(baseOutputPath);
            string extension = Path.GetExtension(baseOutputPath);

            string pricedPath = Path.Combine(directory, $"{fileNameWithoutExt}_PRICED{extension}");
            string unpricedPath = Path.Combine(directory, $"{fileNameWithoutExt}_UNPRICED{extension}");

            // Generate priced version
            string pricedSuffix = GenerateRFQExcel(templatePath, pricedPath, rfq, items, templateId, isPriced: true);

            // Generate unpriced version
            string unpricedSuffix = GenerateRFQExcel(templatePath, unpricedPath, rfq, items, templateId, isPriced: false);

            return (pricedPath, unpricedPath, pricedSuffix, unpricedSuffix);
        }

        /// <summary>
        /// Detects whether the template uses "COMMERCIAL/TECHNICAL" or "PRICED/UNPRICED" terminology
        /// Returns: ("COMMERCIAL", "TECHNICAL") or ("PRICED", "UNPRICED")
        /// </summary>
        private (string pricedTerm, string unpricedTerm) DetectTemplateTerminology(IXLWorksheet worksheet)
        {
            var usedCells = worksheet.CellsUsed();

            foreach (var cell in usedCells)
            {
                if (cell.IsEmpty())
                    continue;

                string cellValue = cell.GetString();

                // Check if template uses "COMMERCIAL" terminology
                if (cellValue.IndexOf("COMMERCIAL", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return ("COMMERCIAL", "TECHNICAL");
                }

                // Check if template uses "PRICED" terminology
                if (cellValue.IndexOf("PRICED", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return ("PRICED", "UNPRICED");
                }
            }

            // Default to PRICED/UNPRICED if neither found
            return ("PRICED", "UNPRICED");
        }

        /// <summary>
        /// Converts PRICED to UNPRICED and COMMERCIAL to TECHNICAL throughout the worksheet
        /// </summary>
        private void ConvertToUnpricedVersion(IXLWorksheet worksheet, string pricedTerm, string unpricedTerm)
        {
            var usedCells = worksheet.CellsUsed();
            foreach (var cell in usedCells)
            {
                if (cell.IsEmpty())
                    continue;

                string cellValue = cell.GetString();
                bool hasChanged = false;

                // Replace using the detected terminology
                if (cellValue.IndexOf(pricedTerm, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    cellValue = System.Text.RegularExpressions.Regex.Replace(
                        cellValue,
                        pricedTerm,
                        unpricedTerm,
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase
                    );
                    hasChanged = true;
                }

                if (hasChanged)
                {
                    cell.Value = cellValue;
                }
            }
        }

        /// <summary>
        /// Converts the template filename from database to full path
        /// Database stores: "CG TEMPLATE.xlsx"
        /// Returns full path based on application directory
        /// Tries multiple locations to find the template
        /// </summary>
        private string GetFullTemplatePath(string templateFileName)
        {
            string foundPath = null;

            // Try multiple possible locations
            List<string> possiblePaths = new List<string>();

            // 1. Application base directory (bin\Debug or bin\Release)
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            possiblePaths.Add(Path.Combine(appDir, "Templates", templateFileName));

            // 2. Two levels up from bin\Debug (project root during development)
            string projectRoot = Path.GetFullPath(Path.Combine(appDir, @"..\..\"));
            possiblePaths.Add(Path.Combine(projectRoot, "Templates", templateFileName));

            // 3. Current directory
            possiblePaths.Add(Path.Combine(Directory.GetCurrentDirectory(), "Templates", templateFileName));

            // Check each possible path
            foreach (string path in possiblePaths)
            {
                if (File.Exists(path))
                {
                    foundPath = path;
                    break;
                }
            }

            // If still not found, show detailed error
            if (foundPath == null)
            {
                string errorMessage = $"Template file not found: {templateFileName}\n\n";
                errorMessage += "Searched in the following locations:\n";

                for (int i = 0; i < possiblePaths.Count; i++)
                {
                    errorMessage += $"{i + 1}. {possiblePaths[i]}\n";
                    errorMessage += $"   Exists: {File.Exists(possiblePaths[i])}\n\n";
                }

                errorMessage += $"Application Directory: {appDir}\n";
                errorMessage += $"Current Directory: {Directory.GetCurrentDirectory()}\n\n";

                // Check if Templates folder exists anywhere
                string templatesFolder1 = Path.Combine(appDir, "Templates");
                string templatesFolder2 = Path.Combine(projectRoot, "Templates");

                errorMessage += $"Templates folder status:\n";
                errorMessage += $"1. {templatesFolder1}\n   Exists: {Directory.Exists(templatesFolder1)}\n\n";
                errorMessage += $"2. {templatesFolder2}\n   Exists: {Directory.Exists(templatesFolder2)}\n\n";

                // Show what files are in Templates folder if it exists
                if (Directory.Exists(templatesFolder2))
                {
                    errorMessage += $"Files found in {templatesFolder2}:\n";
                    string[] files = Directory.GetFiles(templatesFolder2);
                    if (files.Length > 0)
                    {
                        foreach (string file in files)
                        {
                            errorMessage += $"- {Path.GetFileName(file)}\n";
                        }
                    }
                    else
                    {
                        errorMessage += "(No files found)\n";
                    }
                }
                else if (Directory.Exists(templatesFolder1))
                {
                    errorMessage += $"Files found in {templatesFolder1}:\n";
                    string[] files = Directory.GetFiles(templatesFolder1);
                    if (files.Length > 0)
                    {
                        foreach (string file in files)
                        {
                            errorMessage += $"- {Path.GetFileName(file)}\n";
                        }
                    }
                    else
                    {
                        errorMessage += "(No files found)\n";
                    }
                }

                errorMessage += "\nPlease ensure:\n";
                errorMessage += "1. The Templates folder exists in your project root\n";
                errorMessage += "2. The template file is in the Templates folder\n";
                errorMessage += "3. In Visual Studio, right-click Templates folder → Properties → 'Copy to Output Directory' = 'Copy if newer'\n";
                errorMessage += "4. Rebuild your solution";

                throw new FileNotFoundException(errorMessage);
            }

            return foundPath;
        }

        private IXLCell FindPlaceholderCell(IXLWorksheet worksheet, string placeholder)
        {
            var usedCells = worksheet.CellsUsed();
            foreach (var cell in usedCells)
            {
                if (cell.GetString().Equals(placeholder, StringComparison.OrdinalIgnoreCase))
                {
                    return cell;
                }
            }
            return null;
        }

        private void FillHeaderFields(IXLWorksheet worksheet, RFQ rfq, string clientName, List<RFQItem> items)
        {
            var usedCells = worksheet.CellsUsed();

            // Collect all unique delivery times from items (already in weeks)
            var deliveryWeeks = items
                .Select(item => item.DeliveryTime)
                .Distinct()
                .OrderBy(w => w)
                .ToList();

            // Format: "2, 3 WEEKS" or "1 WEEK" (not "2 WEEKS, 3 WEEKS")
            string headerDeliveryTimeText = "";
            if (deliveryWeeks.Count > 0)
            {
                // Determine if we need singular or plural based on the maximum value
                int maxWeeks = deliveryWeeks.Max();
                string weekText = maxWeeks < 2 ? "WEEK" : "WEEKS";

                // Join numbers with comma, then add WEEK/WEEKS at the end
                headerDeliveryTimeText = string.Join(", ", deliveryWeeks) + " " + weekText;
            }

            foreach (var cell in usedCells)
            {
                if (cell.IsEmpty())
                    continue;

                string cellValue = cell.GetString();

                if (cellValue.Equals(Placeholders.ClientName, StringComparison.OrdinalIgnoreCase))
                {
                    cell.Value = clientName;
                }
                else if (cellValue.Equals(Placeholders.RFQCode, StringComparison.OrdinalIgnoreCase))
                {
                    cell.Value = rfq.RFQCode;
                }
                else if (cellValue.Equals(Placeholders.Date, StringComparison.OrdinalIgnoreCase))
                {
                    cell.Value = rfq.CreatedAt.ToString("dd MMMM yyyy");
                }
                else if (cellValue.Equals(Placeholders.QuoteCode, StringComparison.OrdinalIgnoreCase))
                {
                    cell.Value = rfq.QuoteCode;
                }
                else if (cellValue.Equals(Placeholders.DeliveryTerm, StringComparison.OrdinalIgnoreCase))
                {
                    cell.Value = (rfq.DeliveryTerm ?? "");
                }
                else if (cellValue.Equals(Placeholders.DeliveryPoint, StringComparison.OrdinalIgnoreCase))
                {
                    cell.Value = (rfq.DeliveryPoint ?? "");
                }
                else if (cellValue.Equals(Placeholders.Validity, StringComparison.OrdinalIgnoreCase))
                {
                    cell.Value = (rfq.Validity ?? "");
                }
                else if (cellValue.Equals(Placeholders.HeaderDeliveryTime, StringComparison.OrdinalIgnoreCase))
                {
                    cell.Value = headerDeliveryTimeText;
                }
            }
        }

        private void FillItems(IXLWorksheet worksheet, List<RFQItem> items, IXLCell startCell, string effectiveCurrency, bool isPriced)
        {
            int headerRow = startCell.Address.RowNumber;

            string itemNoCol = FindColumnByPlaceholder(worksheet, headerRow, Placeholders.ItemNo);
            string descCol = FindColumnByPlaceholder(worksheet, headerRow, Placeholders.ItemDesc);
            string deliveryCol = FindColumnByPlaceholder(worksheet, headerRow, Placeholders.DeliveryTime);
            string qtyCol = FindColumnByPlaceholder(worksheet, headerRow, Placeholders.Quantity);
            string priceCol = FindColumnByPlaceholder(worksheet, headerRow, Placeholders.UnitPrice);
            string totalCol = FindColumnByPlaceholder(worksheet, headerRow, Placeholders.TotalPrice);

            // Check if template has delivery_time column
            bool hasDeliveryColumn = deliveryCol != null;

            int startRow = headerRow;

            // Calculate total rows needed
            int totalRowsNeeded = 0;
            foreach (var item in items)
            {
                string[] descLines = (item.ItemDesc ?? "").Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

                if (hasDeliveryColumn)
                {
                    // Template has delivery column, no extra row needed under description
                    totalRowsNeeded += descLines.Length + 1; // +1 for spacing between items
                }
                else
                {
                    // No delivery column, add extra row for delivery time under description
                    totalRowsNeeded += descLines.Length + 1 + 1; // +1 for delivery time row, +1 for spacing
                }
            }

            int templateRows = 1;
            int rowsToInsert = totalRowsNeeded - templateRows;

            if (rowsToInsert > 0)
            {
                // Get merged cells info BEFORE inserting rows
                var mergedCellsInTemplateRow = GetMergedCellsInRow(worksheet, startRow);

                worksheet.Row(startRow + templateRows).InsertRowsAbove(rowsToInsert);

                // Copy formatting AND merged cells
                for (int r = 0; r < rowsToInsert; r++)
                {
                    int targetRow = startRow + templateRows + r;
                    worksheet.Row(targetRow).Height = worksheet.Row(startRow).Height;

                    // Copy cell styles and replace currency in the process
                    foreach (var cell in worksheet.Row(startRow).Cells())
                    {
                        if (!cell.IsEmpty() || cell.Style.Fill.BackgroundColor.HasValue)
                        {
                            var targetCell = worksheet.Cell(targetRow, cell.Address.ColumnNumber);
                            targetCell.Style = cell.Style;

                            // If the template cell has a value with "RM", copy it with currency replaced
                            if (!cell.IsEmpty())
                            {
                                string cellValue = cell.GetString();
                                if (cellValue.Contains("RM") && effectiveCurrency != "RM")
                                {
                                    targetCell.Value = cellValue.Replace("RM", effectiveCurrency);
                                }
                                else if (!string.IsNullOrWhiteSpace(cellValue))
                                {
                                    // Copy the value if it's not a placeholder (placeholders will be filled later)
                                    bool isPlaceholder = cellValue.Equals(Placeholders.ItemNo, StringComparison.OrdinalIgnoreCase) ||
                                                        cellValue.Equals(Placeholders.ItemDesc, StringComparison.OrdinalIgnoreCase) ||
                                                        cellValue.Equals(Placeholders.DeliveryTime, StringComparison.OrdinalIgnoreCase) ||
                                                        cellValue.Equals(Placeholders.Quantity, StringComparison.OrdinalIgnoreCase) ||
                                                        cellValue.Equals(Placeholders.UnitPrice, StringComparison.OrdinalIgnoreCase) ||
                                                        cellValue.Equals(Placeholders.TotalPrice, StringComparison.OrdinalIgnoreCase);

                                    if (!isPlaceholder)
                                    {
                                        targetCell.Value = cellValue;
                                    }
                                }
                            }

                            // Also replace currency in number format if needed
                            string numberFormat = cell.Style.NumberFormat.Format;
                            if (!string.IsNullOrEmpty(numberFormat) && numberFormat.Contains("RM") && effectiveCurrency != "RM")
                            {
                                targetCell.Style.NumberFormat.Format = numberFormat.Replace("RM", effectiveCurrency);
                            }
                        }
                    }

                    // Copy merged cells
                    foreach (var mergedRange in mergedCellsInTemplateRow)
                    {
                        int startCol = mergedRange.Item1;
                        int endCol = mergedRange.Item2;

                        var newRange = worksheet.Range(targetRow, startCol, targetRow, endCol);
                        newRange.Merge();
                    }
                }
            }

            // Fill each item
            int currentRow = startRow;
            foreach (var item in items)
            {
                string[] descLines = (item.ItemDesc ?? "").Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                int weeks = item.DeliveryTime; // Already in weeks
                string weekText = weeks < 2 ? "WEEK" : "WEEKS";

                if (itemNoCol != null)
                {
                    worksheet.Cell($"{itemNoCol}{currentRow}").Value = item.ItemNo;
                }

                // Fill description lines
                if (descCol != null)
                {
                    for (int i = 0; i < descLines.Length; i++)
                    {
                        worksheet.Cell($"{descCol}{currentRow + i}").Value = descLines[i];
                        worksheet.Cell($"{descCol}{currentRow + i}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    }

                    // Only add delivery time under description if template has NO delivery_time column
                    if (!hasDeliveryColumn)
                    {
                        int deliveryTextRow = currentRow + descLines.Length;
                        worksheet.Cell($"{descCol}{deliveryTextRow}").Value = $"(Delivery: {weeks} {weekText})";
                        worksheet.Cell($"{descCol}{deliveryTextRow}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                        worksheet.Cell($"{descCol}{deliveryTextRow}").Style.Font.Italic = true;
                    }
                }

                // If template has delivery_time column, fill it
                if (hasDeliveryColumn)
                {
                    worksheet.Cell($"{deliveryCol}{currentRow}").Value = weeks + " " + weekText;
                }

                if (qtyCol != null)
                {
                    worksheet.Cell($"{qtyCol}{currentRow}").Value = item.Quantity + " " + item.UnitName;
                }

                // Handle pricing based on isPriced flag
                if (priceCol != null)
                {
                    if (isPriced)
                    {
                        // Priced version: show actual prices
                        worksheet.Cell($"{priceCol}{currentRow}").Value = item.UnitPrice;
                        worksheet.Cell($"{priceCol}{currentRow}").Style.NumberFormat.Format = "#,##0.00";
                    }
                    else
                    {
                        // Unpriced version: show "Quoted" or "Unquoted"
                        string priceText = item.UnitPrice > 0 ? "Quoted" : "Unquoted";
                        worksheet.Cell($"{priceCol}{currentRow}").Value = priceText;
                        // Clear number formatting for text
                        worksheet.Cell($"{priceCol}{currentRow}").Style.NumberFormat.Format = "@";
                    }
                }

                if (totalCol != null && priceCol != null)
                {
                    if (isPriced)
                    {
                        // Priced version: calculate total
                        worksheet.Cell($"{totalCol}{currentRow}").FormulaA1 = $"={priceCol}{currentRow}*{item.Quantity}";
                        worksheet.Cell($"{totalCol}{currentRow}").Style.NumberFormat.Format = "#,##0.00";
                    }
                    else
                    {
                        // Unpriced version: show "Quoted" or "Unquoted"
                        string totalText = item.UnitPrice > 0 ? "Quoted" : "Unquoted";
                        worksheet.Cell($"{totalCol}{currentRow}").Value = totalText;
                        // Clear number formatting for text
                        worksheet.Cell($"{totalCol}{currentRow}").Style.NumberFormat.Format = "@";
                    }
                }

                // Move to next item
                if (hasDeliveryColumn)
                {
                    // Template has delivery column: description lines + spacing
                    currentRow += descLines.Length + 1;
                }
                else
                {
                    // No delivery column: description lines + delivery line + spacing
                    currentRow += descLines.Length + 1 + 1;
                }
            }
        }

        /// <summary>
        /// Get all merged cell ranges in a specific row
        /// Returns list of (startColumn, endColumn) tuples
        /// </summary>
        private List<Tuple<int, int>> GetMergedCellsInRow(IXLWorksheet worksheet, int rowNumber)
        {
            var mergedRanges = new List<Tuple<int, int>>();

            foreach (var mergedRange in worksheet.MergedRanges)
            {
                var rangeAddress = mergedRange.RangeAddress;

                // Check if this merged range includes our row
                if (rangeAddress.FirstAddress.RowNumber <= rowNumber &&
                    rangeAddress.LastAddress.RowNumber >= rowNumber)
                {
                    mergedRanges.Add(new Tuple<int, int>(
                        rangeAddress.FirstAddress.ColumnNumber,
                        rangeAddress.LastAddress.ColumnNumber
                    ));
                }
            }

            return mergedRanges;
        }

        private string FindColumnByPlaceholder(IXLWorksheet worksheet, int row, string placeholder)
        {
            if (string.IsNullOrEmpty(placeholder))
                return null;

            var rowCells = worksheet.Row(row).Cells();
            foreach (var cell in rowCells)
            {
                if (cell.GetString().Equals(placeholder, StringComparison.OrdinalIgnoreCase))
                {
                    return cell.Address.ColumnLetter;
                }
            }
            return null;
        }

        private void FillSummary(IXLWorksheet worksheet, List<RFQItem> items, decimal discount, bool isPriced)
        {
            decimal subtotal = items.Sum(item => item.UnitPrice * item.Quantity);

            // Calculate discount amount and total price
            // If discount is 0, no discount is applied
            decimal discountAmount = (discount > 0) ? (subtotal * discount / 100) : 0;
            decimal totalPrice = subtotal - discountAmount;

            var usedCells = worksheet.CellsUsed();

            foreach (var cell in usedCells)
            {
                if (cell.IsEmpty())
                    continue;

                string cellValue = cell.GetString();

                if (cellValue.Equals(Placeholders.Subtotal, StringComparison.OrdinalIgnoreCase))
                {
                    if (isPriced)
                    {
                        cell.Value = subtotal;
                        cell.Style.NumberFormat.Format = "#,##0.00";
                    }
                    else
                    {
                        cell.Value = "Quoted";
                        cell.Style.NumberFormat.Format = "@";
                    }

                }
                else if (cellValue.Equals(Placeholders.Discount, StringComparison.OrdinalIgnoreCase))
                {
                    if (isPriced)
                    {
                        cell.Value = discount / 100;
                        cell.Style.NumberFormat.Format = "#,##0.00 %";
                    }
                    else
                    {
                        // For unpriced, discount is always 0
                        cell.Value = 0;
                        cell.Style.NumberFormat.Format = "#,##0.00 %";
                    }
                }
                else if (cellValue.Equals(Placeholders.SummaryTotal, StringComparison.OrdinalIgnoreCase))
                {
                    if (isPriced)
                    {
                        cell.Value = totalPrice;
                        cell.Style.NumberFormat.Format = "#,##0.00";
                    }
                    else
                    {
                        cell.Value = "Quoted";
                        cell.Style.NumberFormat.Format = "@";
                    }
                }

            }
        }

        /// <summary>
        /// Replaces all instances of "RM" with the selected currency throughout the worksheet.
        /// This includes both cell values and number formats.
        /// </summary>
        private void ReplaceCurrencyInWorksheet(IXLWorksheet worksheet, string currency)
        {
            if (string.IsNullOrEmpty(currency) || currency == "RM")
                return; // No replacement needed if currency is RM or null

            var usedCells = worksheet.CellsUsed();
            foreach (var cell in usedCells)
            {
                if (cell.IsEmpty())
                    continue;

                // Replace "RM" in cell text values
                string cellValue = cell.GetString();
                if (cellValue.Contains("RM"))
                {
                    cell.Value = cellValue.Replace("RM", currency);
                }

                // Replace "RM" in number formats
                // Example: "RM #,##0.00" becomes "USD #,##0.00"
                string numberFormat = cell.Style.NumberFormat.Format;
                if (!string.IsNullOrEmpty(numberFormat) && numberFormat.Contains("RM"))
                {
                    cell.Style.NumberFormat.Format = numberFormat.Replace("RM", currency);
                }
            }
        }

        /// <summary>
        /// Apply bottom borders to page breaks
        /// Works for multiple templates with configurable first page break row, row increment, and column range
        /// </summary>
        private void ApplyBottomBordersAtPageBreaks(IXLWorksheet worksheet, string templatePath)
        {
            try
            {
                string templateFileName = Path.GetFileName(templatePath);

                // Template configuration: Dictionary of template name -> (firstPageBreakRow, rowIncrement, startColumn, endColumn)
                var templateConfig = new Dictionary<string, (int firstRow, int increment, string startCol, string endCol)>(StringComparer.OrdinalIgnoreCase)
                {
                    { "CG TEMPLATE.xlsx", (68, 43, "A", "H") },
                    { "DE TEMPLATE.xlsx", (61, 43, "A", "T") },
                    { "GA TEMPLATE.xlsx", (65, 41, "A", "H") },
                    { "OP TEMPLATE.xlsx", (70, 46, "A", "H") },
                    { "PO TEMPLATE.xlsx", (70, 66, "A", "J") },
                    { "SC TEMPLATE.xlsx", (66, 50, "A", "F") },
                };

                // Check if this template is configured
                if (!templateConfig.ContainsKey(templateFileName))
                {
                    return; // Skip for templates not in configuration
                }

                // Get template configuration
                var (firstPageBreakRow, rowIncrement, startColumn, endColumn) = templateConfig[templateFileName];

                // Check if worksheet has more than one page
                if (!HasMultiplePages(worksheet))
                {
                    return; // Only apply borders if there are page breaks
                }

                // Get the used range to know where data ends
                var usedRange = worksheet.RangeUsed();
                if (usedRange == null)
                    return;

                int lastRow = usedRange.LastRow().RowNumber();

                // Calculate page break rows based on configuration
                int pageBreakRow = firstPageBreakRow;

                while (pageBreakRow <= lastRow)
                {
                    // Apply thick border to this row with configured column range
                    // (even if row is empty - page break rows should always have border)
                    ApplyThickBorderToRange(worksheet, pageBreakRow, startColumn, endColumn);
                    System.Diagnostics.Debug.WriteLine($"Applied border to row {pageBreakRow} ({templateFileName})");

                    // Next page break
                    pageBreakRow += rowIncrement;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error applying bottom borders at page breaks: {ex.Message}");
            }
        }

        /// <summary>
        /// Check if the worksheet will have multiple pages when printed
        /// </summary>
        private bool HasMultiplePages(IXLWorksheet worksheet)
        {
            try
            {
                // Get the used range
                var usedRange = worksheet.RangeUsed();
                if (usedRange == null)
                    return false;

                int lastRow = usedRange.LastRow().RowNumber();

                // Get page setup
                var pageSetup = worksheet.PageSetup;
                double topMargin = pageSetup.Margins.Top;
                double bottomMargin = pageSetup.Margins.Bottom;

                double pageHeight = 842; // A4 height in points
                double availableHeight = pageHeight - topMargin - bottomMargin;

                // Calculate total height
                double totalHeight = 0;
                for (int row = 1; row <= lastRow; row++)
                {
                    var rowHeight = worksheet.Row(row).Height;
                    if (rowHeight > 0)
                    {
                        totalHeight += rowHeight;
                    }
                    else
                    {
                        totalHeight += 15; // Default row height
                    }
                }

                // If total height exceeds available height, there are multiple pages
                return totalHeight > availableHeight;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Apply thick bottom border to a range of cells in a specific row
        /// </summary>
        private void ApplyThickBorderToRange(IXLWorksheet worksheet, int rowNumber, string startColumn, string endColumn)
        {
            try
            {
                var range = worksheet.Range($"{startColumn}{rowNumber}:{endColumn}{rowNumber}");
                foreach (var cell in range.Cells())
                {
                    cell.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                    cell.Style.Border.BottomBorderColor = XLColor.Black;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error applying border to range: {ex.Message}");
            }
        }
    }
}