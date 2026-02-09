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

        public void GenerateRFQExcel(string templatePath, string outputPath, RFQ rfq, List<RFQItem> items, int templateId)
        {
            // Convert database path (filename only) to full path
            string fullTemplatePath = GetFullTemplatePath(templatePath);

            using (var workbook = new XLWorkbook(fullTemplatePath))
            {
                var worksheet = workbook.Worksheet(1);

                var client = clientRepo.GetClientById(rfq.ClientId);
                string clientName = client?.ClientName ?? "";

                var itemStartCell = FindPlaceholderCell(worksheet, Placeholders.ItemNo);

                // Fill header fields including header_delivery_time placeholder
                FillHeaderFields(worksheet, rfq, clientName, items);

                if (itemStartCell != null)
                {
                    FillItems(worksheet, items, itemStartCell);
                }

                FillSummary(worksheet, items, rfq.Discount);

                workbook.SaveAs(outputPath);
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

        private void FillItems(IXLWorksheet worksheet, List<RFQItem> items, IXLCell startCell)
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

                    // Copy cell styles
                    foreach (var cell in worksheet.Row(startRow).Cells())
                    {
                        if (!cell.IsEmpty() || cell.Style.Fill.BackgroundColor.HasValue)
                        {
                            worksheet.Cell(targetRow, cell.Address.ColumnNumber).Style = cell.Style;
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

                if (priceCol != null)
                {
                    worksheet.Cell($"{priceCol}{currentRow}").Value = item.UnitPrice;
                    worksheet.Cell($"{priceCol}{currentRow}").Style.NumberFormat.Format = "#,##0.00";
                }

                if (totalCol != null && priceCol != null)
                {
                    worksheet.Cell($"{totalCol}{currentRow}").FormulaA1 = $"={priceCol}{currentRow}*{item.Quantity}";
                    worksheet.Cell($"{totalCol}{currentRow}").Style.NumberFormat.Format = "#,##0.00";
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

        private void FillSummary(IXLWorksheet worksheet, List<RFQItem> items, decimal discount)
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
                    cell.Value = subtotal;
                    cell.Style.NumberFormat.Format = "#,##0.00";
                }
                else if (cellValue.Equals(Placeholders.Discount, StringComparison.OrdinalIgnoreCase))
                {
                    cell.Value = discount / 100;
                    cell.Style.NumberFormat.Format = "#,##0.00 %";
                }
                else if (cellValue.Equals(Placeholders.SummaryTotal, StringComparison.OrdinalIgnoreCase))
                {
                    cell.Value = totalPrice;
                    cell.Style.NumberFormat.Format = "#,##0.00";
                }
            }
        }
    }
}