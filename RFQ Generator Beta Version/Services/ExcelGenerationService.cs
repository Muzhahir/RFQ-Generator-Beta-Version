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

        public string GenerateRFQExcel(string templatePath, string outputPath, RFQ rfq, List<RFQItem> items, int templateId)
        {
            return GenerateRFQExcel(templatePath, outputPath, rfq, items, templateId, isPriced: true);
        }

        public string GenerateRFQExcel(string templatePath, string outputPath, RFQ rfq, List<RFQItem> items, int templateId, bool isPriced)
        {
            string fullTemplatePath = GetFullTemplatePath(templatePath);

            using (var workbook = new XLWorkbook(fullTemplatePath))
            {
                var worksheet = workbook.Worksheet(1);

                var (pricedTerm, unpricedTerm) = DetectTemplateTerminology(worksheet);
                string versionSuffix = isPriced ? pricedTerm : unpricedTerm;

                var client = clientRepo.GetClientById(rfq.ClientId);
                string clientName = client?.ClientName ?? "";

                var itemStartCell = FindPlaceholderCell(worksheet, Placeholders.ItemNo);

                string effectiveCurrency = rfq.Currency ?? "RM";

                System.Diagnostics.Debug.WriteLine($"[ExcelService] Template terminology: {pricedTerm}/{unpricedTerm}");
                System.Diagnostics.Debug.WriteLine($"[ExcelService] Version suffix: {versionSuffix}");
                System.Diagnostics.Debug.WriteLine($"[ExcelService] Currency: {effectiveCurrency}");
                System.Diagnostics.Debug.WriteLine($"[ExcelService] Is Priced: {isPriced}");

                if (!isPriced)
                    ConvertToUnpricedVersion(worksheet, pricedTerm, unpricedTerm);

                FillHeaderFields(worksheet, rfq, clientName, items);

                if (itemStartCell != null)
                    FillItems(worksheet, items, itemStartCell, effectiveCurrency, isPriced);

                decimal effectiveDiscount = isPriced ? rfq.Discount : 0;

                var (subtotalRow, summaryTotalRow) = FillSummary(worksheet, items, effectiveDiscount, isPriced);

                var (insertedFrom, insertedTo) = EnsureTAndCFitsOnPage(worksheet, fullTemplatePath, subtotalRow, summaryTotalRow);

                ReplaceCurrencyInWorksheet(worksheet, effectiveCurrency);

                ApplyBottomBordersAtPageBreaks(worksheet, fullTemplatePath, insertedFrom, insertedTo);

                workbook.SaveAs(outputPath);

                return versionSuffix;
            }
        }

        public (string pricedPath, string unpricedPath, string pricedSuffix, string unpricedSuffix) GenerateBothVersionsExcel(
            string templatePath,
            string baseOutputPath,
            RFQ rfq,
            List<RFQItem> items,
            int templateId)
        {
            string directory = Path.GetDirectoryName(baseOutputPath);
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(baseOutputPath);
            string extension = Path.GetExtension(baseOutputPath);

            string tempPricedPath = Path.Combine(directory, $"{fileNameWithoutExt}_TEMP_PRICED{extension}");
            string tempUnpricedPath = Path.Combine(directory, $"{fileNameWithoutExt}_TEMP_UNPRICED{extension}");

            string pricedSuffix = GenerateRFQExcel(templatePath, tempPricedPath, rfq, items, templateId, isPriced: true);
            string unpricedSuffix = GenerateRFQExcel(templatePath, tempUnpricedPath, rfq, items, templateId, isPriced: false);

            string pricedPath = Path.Combine(directory, $"{fileNameWithoutExt}_{pricedSuffix}{extension}");
            string unpricedPath = Path.Combine(directory, $"{fileNameWithoutExt}_{unpricedSuffix}{extension}");

            if (File.Exists(pricedPath)) File.Delete(pricedPath);
            if (File.Exists(unpricedPath)) File.Delete(unpricedPath);

            File.Move(tempPricedPath, pricedPath);
            File.Move(tempUnpricedPath, unpricedPath);

            return (pricedPath, unpricedPath, pricedSuffix, unpricedSuffix);
        }

        private (string pricedTerm, string unpricedTerm) DetectTemplateTerminology(IXLWorksheet worksheet)
        {
            foreach (var cell in worksheet.CellsUsed())
            {
                if (cell.IsEmpty()) continue;
                string cellValue = cell.GetString();

                if (cellValue.IndexOf("COMMERCIAL", StringComparison.OrdinalIgnoreCase) >= 0)
                    return ("COMMERCIAL", "TECHNICAL");

                if (cellValue.IndexOf("PRICED", StringComparison.OrdinalIgnoreCase) >= 0)
                    return ("PRICED", "UNPRICED");
            }

            return ("PRICED", "UNPRICED");
        }

        private void ConvertToUnpricedVersion(IXLWorksheet worksheet, string pricedTerm, string unpricedTerm)
        {
            foreach (var cell in worksheet.CellsUsed())
            {
                if (cell.IsEmpty()) continue;
                string cellValue = cell.GetString();

                if (cellValue.IndexOf(pricedTerm, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    cell.Value = System.Text.RegularExpressions.Regex.Replace(
                        cellValue, pricedTerm, unpricedTerm,
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                }
            }
        }

        private string GetFullTemplatePath(string templateFileName)
        {
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            string projectRoot = Path.GetFullPath(Path.Combine(appDir, @"..\..\"));

            var possiblePaths = new List<string>
            {
                Path.Combine(appDir,      "Templates", templateFileName),
                Path.Combine(projectRoot, "Templates", templateFileName),
                Path.Combine(Directory.GetCurrentDirectory(), "Templates", templateFileName),
            };

            foreach (string path in possiblePaths)
                if (File.Exists(path)) return path;

            string errorMessage = $"Template file not found: {templateFileName}\n\nSearched in:\n";
            for (int i = 0; i < possiblePaths.Count; i++)
                errorMessage += $"{i + 1}. {possiblePaths[i]}\n   Exists: {File.Exists(possiblePaths[i])}\n\n";

            string templatesFolder = Path.Combine(projectRoot, "Templates");
            if (Directory.Exists(templatesFolder))
            {
                errorMessage += $"Files found in {templatesFolder}:\n";
                foreach (string f in Directory.GetFiles(templatesFolder))
                    errorMessage += $"- {Path.GetFileName(f)}\n";
            }

            errorMessage += "\nPlease ensure:\n";
            errorMessage += "1. The Templates folder exists in your project root\n";
            errorMessage += "2. The template file is in the Templates folder\n";
            errorMessage += "3. In Visual Studio: right-click Templates folder → Properties → 'Copy to Output Directory' = 'Copy if newer'\n";
            errorMessage += "4. Rebuild your solution";

            throw new FileNotFoundException(errorMessage);
        }

        private IXLCell FindPlaceholderCell(IXLWorksheet worksheet, string placeholder)
        {
            foreach (var cell in worksheet.CellsUsed())
                if (cell.GetString().Equals(placeholder, StringComparison.OrdinalIgnoreCase))
                    return cell;
            return null;
        }

        private void FillHeaderFields(IXLWorksheet worksheet, RFQ rfq, string clientName, List<RFQItem> items)
        {
            var deliveryWeeks = items
                .Select(i => i.DeliveryTime)
                .Distinct()
                .OrderBy(w => w)
                .ToList();

            string headerDeliveryTimeText = "";
            if (deliveryWeeks.Count > 0)
            {
                int maxWeeks = deliveryWeeks.Max();
                string weekText = maxWeeks < 2 ? "WEEK" : "WEEKS";
                headerDeliveryTimeText = string.Join(", ", deliveryWeeks) + " " + weekText;
            }

            foreach (var cell in worksheet.CellsUsed())
            {
                if (cell.IsEmpty()) continue;
                string cellValue = cell.GetString();

                if (cellValue.Equals(Placeholders.ClientName, StringComparison.OrdinalIgnoreCase)) cell.Value = clientName;
                else if (cellValue.Equals(Placeholders.RFQCode, StringComparison.OrdinalIgnoreCase)) cell.Value = rfq.RFQCode;
                else if (cellValue.Equals(Placeholders.Date, StringComparison.OrdinalIgnoreCase)) cell.Value = rfq.CreatedAt.ToString("dd MMMM yyyy");
                else if (cellValue.Equals(Placeholders.QuoteCode, StringComparison.OrdinalIgnoreCase)) cell.Value = rfq.QuoteCode;
                else if (cellValue.Equals(Placeholders.DeliveryTerm, StringComparison.OrdinalIgnoreCase)) cell.Value = rfq.DeliveryTerm ?? "";
                else if (cellValue.Equals(Placeholders.DeliveryPoint, StringComparison.OrdinalIgnoreCase)) cell.Value = rfq.DeliveryPoint ?? "";
                else if (cellValue.Equals(Placeholders.Validity, StringComparison.OrdinalIgnoreCase)) cell.Value = rfq.Validity ?? "";
                else if (cellValue.Equals(Placeholders.HeaderDeliveryTime, StringComparison.OrdinalIgnoreCase)) cell.Value = headerDeliveryTimeText;
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

            bool hasDeliveryColumn = deliveryCol != null;
            int startRow = headerRow;

            int totalRowsNeeded = 0;
            foreach (var item in items)
            {
                string[] descLines = (item.ItemDesc ?? "").Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                totalRowsNeeded += descLines.Length + 1;
                if (!hasDeliveryColumn) totalRowsNeeded += 1;
            }

            int rowsToInsert = totalRowsNeeded - 1;

            if (rowsToInsert > 0)
            {
                var mergedCellsInTemplateRow = GetMergedCellsInRow(worksheet, startRow);
                worksheet.Row(startRow + 1).InsertRowsAbove(rowsToInsert);

                for (int r = 0; r < rowsToInsert; r++)
                {
                    int targetRow = startRow + 1 + r;
                    worksheet.Row(targetRow).Height = worksheet.Row(startRow).Height;

                    foreach (var cell in worksheet.Row(startRow).Cells())
                    {
                        if (!cell.IsEmpty() || cell.Style.Fill.BackgroundColor.HasValue)
                        {
                            var targetCell = worksheet.Cell(targetRow, cell.Address.ColumnNumber);
                            targetCell.Style = cell.Style;

                            if (!cell.IsEmpty())
                            {
                                string cv = cell.GetString();
                                bool isPlaceholder =
                                    cv.Equals(Placeholders.ItemNo, StringComparison.OrdinalIgnoreCase) ||
                                    cv.Equals(Placeholders.ItemDesc, StringComparison.OrdinalIgnoreCase) ||
                                    cv.Equals(Placeholders.DeliveryTime, StringComparison.OrdinalIgnoreCase) ||
                                    cv.Equals(Placeholders.Quantity, StringComparison.OrdinalIgnoreCase) ||
                                    cv.Equals(Placeholders.UnitPrice, StringComparison.OrdinalIgnoreCase) ||
                                    cv.Equals(Placeholders.TotalPrice, StringComparison.OrdinalIgnoreCase);

                                if (!isPlaceholder)
                                {
                                    targetCell.Value = cv.Contains("RM") && effectiveCurrency != "RM"
                                        ? cv.Replace("RM", effectiveCurrency)
                                        : cv;
                                }
                            }

                            string numberFormat = cell.Style.NumberFormat.Format;
                            if (!string.IsNullOrEmpty(numberFormat) && numberFormat.Contains("RM") && effectiveCurrency != "RM")
                                targetCell.Style.NumberFormat.Format = numberFormat.Replace("RM", effectiveCurrency);
                        }
                    }

                    foreach (var mergedRange in mergedCellsInTemplateRow)
                        worksheet.Range(targetRow, mergedRange.Item1, targetRow, mergedRange.Item2).Merge();
                }
            }

            int currentRow = startRow;

            foreach (var item in items)
            {
                string[] descLines = (item.ItemDesc ?? "").Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                int weeks = item.DeliveryTime;
                string weekText = weeks < 2 ? "WEEK" : "WEEKS";

                if (itemNoCol != null)
                    worksheet.Cell($"{itemNoCol}{currentRow}").Value = item.ItemNo;

                if (descCol != null)
                {
                    for (int i = 0; i < descLines.Length; i++)
                    {
                        worksheet.Cell($"{descCol}{currentRow + i}").Value = descLines[i];
                        worksheet.Cell($"{descCol}{currentRow + i}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    }

                    if (!hasDeliveryColumn)
                    {
                        int deliveryTextRow = currentRow + descLines.Length;
                        worksheet.Cell($"{descCol}{deliveryTextRow}").Value = $"(Delivery: {weeks} {weekText})";
                        worksheet.Cell($"{descCol}{deliveryTextRow}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                        worksheet.Cell($"{descCol}{deliveryTextRow}").Style.Font.Italic = true;
                    }
                }

                if (hasDeliveryColumn)
                    worksheet.Cell($"{deliveryCol}{currentRow}").Value = $"{weeks} {weekText}";

                if (qtyCol != null)
                    worksheet.Cell($"{qtyCol}{currentRow}").Value = $"{item.Quantity} {item.UnitName}";

                if (priceCol != null)
                {
                    if (isPriced)
                    {
                        worksheet.Cell($"{priceCol}{currentRow}").Value = item.UnitPrice;
                        worksheet.Cell($"{priceCol}{currentRow}").Style.NumberFormat.Format = "#,##0.00";
                    }
                    else
                    {
                        worksheet.Cell($"{priceCol}{currentRow}").Value = item.UnitPrice > 0 ? "Quoted" : "Unquoted";
                        worksheet.Cell($"{priceCol}{currentRow}").Style.NumberFormat.Format = "@";
                    }
                }

                if (totalCol != null && priceCol != null)
                {
                    if (isPriced)
                    {
                        worksheet.Cell($"{totalCol}{currentRow}").FormulaA1 = $"={priceCol}{currentRow}*{item.Quantity}";
                        worksheet.Cell($"{totalCol}{currentRow}").Style.NumberFormat.Format = "#,##0.00";
                    }
                    else
                    {
                        worksheet.Cell($"{totalCol}{currentRow}").Value = item.UnitPrice > 0 ? "Quoted" : "Unquoted";
                        worksheet.Cell($"{totalCol}{currentRow}").Style.NumberFormat.Format = "@";
                    }
                }

                currentRow += hasDeliveryColumn
                    ? descLines.Length + 1
                    : descLines.Length + 1 + 1;
            }
        }

        private List<Tuple<int, int>> GetMergedCellsInRow(IXLWorksheet worksheet, int rowNumber)
        {
            var mergedRanges = new List<Tuple<int, int>>();
            foreach (var mergedRange in worksheet.MergedRanges)
            {
                var addr = mergedRange.RangeAddress;
                if (addr.FirstAddress.RowNumber <= rowNumber && addr.LastAddress.RowNumber >= rowNumber)
                    mergedRanges.Add(new Tuple<int, int>(addr.FirstAddress.ColumnNumber, addr.LastAddress.ColumnNumber));
            }
            return mergedRanges;
        }

        private string FindColumnByPlaceholder(IXLWorksheet worksheet, int row, string placeholder)
        {
            if (string.IsNullOrEmpty(placeholder)) return null;
            foreach (var cell in worksheet.Row(row).Cells())
                if (cell.GetString().Equals(placeholder, StringComparison.OrdinalIgnoreCase))
                    return cell.Address.ColumnLetter;
            return null;
        }

        private (int subtotalRow, int summaryTotalRow) FillSummary(IXLWorksheet worksheet, List<RFQItem> items, decimal discount, bool isPriced)
        {
            decimal subtotal = items.Sum(i => i.UnitPrice * i.Quantity);
            decimal discountAmount = discount > 0 ? subtotal * discount / 100 : 0;
            decimal totalPrice = subtotal - discountAmount;
            int subtotalRow = -1;
            int summaryTotalRow = -1;

            foreach (var cell in worksheet.CellsUsed())
            {
                if (cell.IsEmpty()) continue;
                string cellValue = cell.GetString();

                if (cellValue.Equals(Placeholders.Subtotal, StringComparison.OrdinalIgnoreCase))
                {
                    subtotalRow = cell.Address.RowNumber;
                    if (isPriced) { cell.Value = subtotal; cell.Style.NumberFormat.Format = "#,##0.00"; }
                    else { cell.Value = "Quoted"; cell.Style.NumberFormat.Format = "@"; }
                }
                else if (cellValue.Equals(Placeholders.Discount, StringComparison.OrdinalIgnoreCase))
                {
                    if (isPriced) { cell.Value = discount / 100; cell.Style.NumberFormat.Format = "#,##0.00 %"; }
                    else { cell.Value = 0; cell.Style.NumberFormat.Format = "#,##0.00 %"; }
                }
                else if (cellValue.Equals(Placeholders.SummaryTotal, StringComparison.OrdinalIgnoreCase))
                {
                    summaryTotalRow = cell.Address.RowNumber;
                    if (isPriced) { cell.Value = totalPrice; cell.Style.NumberFormat.Format = "#,##0.00"; }
                    else { cell.Value = "Quoted"; cell.Style.NumberFormat.Format = "@"; }
                }
            }

            System.Diagnostics.Debug.WriteLine($"[FillSummary] subtotalRow={subtotalRow}, summaryTotalRow={summaryTotalRow}");
            return (subtotalRow, summaryTotalRow);
        }

        private (int insertedFrom, int insertedTo) EnsureTAndCFitsOnPage(IXLWorksheet worksheet, string templatePath, int subtotalRow, int summaryTotalRow)
        {
            try
            {
                if (subtotalRow < 0) return (-1, -1);

                string templateFileName = Path.GetFileName(templatePath);

                var combinedBlockSizes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
                {
                    { "CG TEMPLATE.xlsx",   19 },
                    { "DE TEMPLATE.xlsx",   24 },
                    { "GA TEMPLATE.xlsx",   26 },
                    { "MA TEMPLATE.xlsx",   26 },
                    { "OGIT TEMPLATE.xlsx", 29 },
                    { "OP TEMPLATE.xlsx",   33 },
                    { "PO TEMPLATE.xlsx",   17 },
                    { "SC TEMPLATE.xlsx",   25 },
                };

                var templateConfig = new Dictionary<string, (int firstRow, int increment)>(StringComparer.OrdinalIgnoreCase)
                {
                    { "CG TEMPLATE.xlsx",   (68, 43) },
                    { "DE TEMPLATE.xlsx",   (63, 43) },
                    { "GA TEMPLATE.xlsx",   (65, 42) },
                    { "MA TEMPLATE.xlsx",   (61, 44) },
                    { "OGIT TEMPLATE.xlsx", (67, 46) },
                    { "OP TEMPLATE.xlsx",   (70, 50) },
                    { "PO TEMPLATE.xlsx",   (70, 66) },
                    { "SC TEMPLATE.xlsx",   (66, 51) },
                };

                if (!templateConfig.ContainsKey(templateFileName)) return (-1, -1);
                if (!combinedBlockSizes.TryGetValue(templateFileName, out int combinedBlockSize)) return (-1, -1);
                if (combinedBlockSize <= 0) return (-1, -1);

                var (firstPageBreakRow, pageIncrement) = templateConfig[templateFileName];

                int blockStart = subtotalRow;
                int blockEnd = subtotalRow + combinedBlockSize - 1;

                int pageBreakRow = firstPageBreakRow;
                while (pageBreakRow < blockStart)
                    pageBreakRow += pageIncrement;

                System.Diagnostics.Debug.WriteLine($"[TAndC] blockStart={blockStart}, blockEnd={blockEnd}, combinedBlockSize={combinedBlockSize}, pageBreakRow={pageBreakRow}");

                if (pageBreakRow > blockEnd)
                {
                    System.Diagnostics.Debug.WriteLine($"[TAndC] No page break inside combined block. Nothing to do.");
                    return (-1, -1);
                }

                if (pageBreakRow == blockStart)
                {
                    System.Diagnostics.Debug.WriteLine($"[TAndC] Block already starts on a new page.");
                    return (-1, -1);
                }

                // +1 so block lands on the first row of the next page, not the last row of the current page
                int rowsToInsert = (pageBreakRow + 1) - blockStart;
                int insertAt = blockStart;
                int sourceStyleRow = insertAt - 1;

                System.Diagnostics.Debug.WriteLine($"[TAndC] Page break at row {pageBreakRow} inside block. Inserting {rowsToInsert} rows at {insertAt}.");

                worksheet.Row(insertAt).InsertRowsAbove(rowsToInsert);

                for (int r = insertAt; r < insertAt + rowsToInsert; r++)
                {
                    worksheet.Row(r).Height = worksheet.Row(sourceStyleRow).Height;

                    foreach (var sourceCell in worksheet.Row(sourceStyleRow).CellsUsed())
                    {
                        var targetCell = worksheet.Cell(r, sourceCell.Address.ColumnNumber);
                        targetCell.Style = sourceCell.Style;
                        targetCell.Value = "";
                    }
                }

                for (int c = 1; c <= 30; c++)
                {
                    for (int r = insertAt; r < insertAt + rowsToInsert; r++)
                    {
                        var cell = worksheet.Cell(r, c);
                        cell.Style.Border.BottomBorder = XLBorderStyleValues.None;
                        cell.Style.Border.TopBorder = XLBorderStyleValues.None;
                        cell.Style.Border.LeftBorder = XLBorderStyleValues.None;
                        cell.Style.Border.RightBorder = XLBorderStyleValues.None;
                    }
                }

                return (insertAt, insertAt + rowsToInsert - 1);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TAndC] Error: {ex.Message}");
            }

            return (-1, -1);
        }

        private void ReplaceCurrencyInWorksheet(IXLWorksheet worksheet, string currency)
        {
            if (string.IsNullOrEmpty(currency) || currency == "RM") return;

            foreach (var cell in worksheet.CellsUsed())
            {
                if (cell.IsEmpty()) continue;

                string cellValue = cell.GetString();
                if (cellValue.Contains("RM"))
                    cell.Value = cellValue.Replace("RM", currency);

                string numberFormat = cell.Style.NumberFormat.Format;
                if (!string.IsNullOrEmpty(numberFormat) && numberFormat.Contains("RM"))
                    cell.Style.NumberFormat.Format = numberFormat.Replace("RM", currency);
            }
        }

        private void ApplyBottomBordersAtPageBreaks(IXLWorksheet worksheet, string templatePath, int insertedFrom, int insertedTo)
        {
            try
            {
                string templateFileName = Path.GetFileName(templatePath);

                var templateConfig = new Dictionary<string, (int firstRow, int increment, string startCol, string endCol)>(StringComparer.OrdinalIgnoreCase)
                {
                    { "CG TEMPLATE.xlsx",   (68, 43, "A", "H") },
                    { "DE TEMPLATE.xlsx",   (61, 43, "A", "T") },
                    { "GA TEMPLATE.xlsx",   (65, 42, "A", "H") },
                    { "MA TEMPLATE.xlsx",   (61, 44, "A", "H") },
                    { "OGIT TEMPLATE.xlsx", (67, 46, "A", "H") },
                    { "OP TEMPLATE.xlsx",   (70, 50, "A", "H") },
                    { "PO TEMPLATE.xlsx",   (70, 66, "A", "I") },
                    { "SC TEMPLATE.xlsx",   (66, 51, "A", "F") },
                };

                if (!templateConfig.ContainsKey(templateFileName)) return;

                var noBorderTemplates = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "MA TEMPLATE.xlsx",
                    "OGIT TEMPLATE.xlsx",
                };

                if (noBorderTemplates.Contains(templateFileName)) return;
                var (firstPageBreakRow, rowIncrement, startColumn, endColumn) = templateConfig[templateFileName];

                var usedRange = worksheet.RangeUsed();
                if (usedRange == null) return;

                int lastRow = usedRange.LastRow().RowNumber();

                int lastItemDescRow = FindLastItemDescriptionRow(worksheet);
                System.Diagnostics.Debug.WriteLine($"[ApplyBorders] Last item description row: {lastItemDescRow}");

                int firstSummaryRow = FindFirstSummaryRow(worksheet);
                System.Diagnostics.Debug.WriteLine($"[ApplyBorders] First summary row: {firstSummaryRow}");

                EnsureSubtotalTopBorder(worksheet, firstSummaryRow, startColumn, endColumn);

                if (!HasMultiplePages(worksheet))
                {
                    System.Diagnostics.Debug.WriteLine($"[ApplyBorders] Single page, no page break borders needed");

                    if (lastItemDescRow > 0)
                    {
                        int tableClosingRow = FindTableClosingBorderRow(worksheet, lastItemDescRow, endColumn);
                        if (tableClosingRow > 0)
                        {
                            ApplyThickBorderToRange(worksheet, tableClosingRow, startColumn, endColumn);
                            System.Diagnostics.Debug.WriteLine($"[ApplyBorders] Single page — border applied at row {tableClosingRow} (table closing row)");
                        }
                    }

                    return;
                }

                System.Diagnostics.Debug.WriteLine($"[ApplyBorders] Multi-page document, checking page break rows");

                int pageBreakRow = firstPageBreakRow;
                while (pageBreakRow <= lastRow)
                {
                    if (insertedFrom >= 0 && pageBreakRow >= insertedFrom && pageBreakRow <= insertedTo)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ApplyBorders] Row {pageBreakRow}: Skipped (inside inserted blank rows)");
                        pageBreakRow += rowIncrement;
                        continue;
                    }

                    if (pageBreakRow <= lastItemDescRow)
                    {
                        if (IsRowInsideItemsTable(worksheet, pageBreakRow, startColumn))
                        {
                            ApplyThickBorderToRange(worksheet, pageBreakRow, startColumn, endColumn);
                            System.Diagnostics.Debug.WriteLine($"[ApplyBorders] Row {pageBreakRow}: Border applied (page break within items)");
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[ApplyBorders] Row {pageBreakRow}: Skipped (not in items table)");
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[ApplyBorders] Row {pageBreakRow}: Skipped (after last item at {lastItemDescRow})");
                    }

                    pageBreakRow += rowIncrement;
                }

                if (lastItemDescRow > 0)
                {
                    int tableClosingRow = FindTableClosingBorderRow(worksheet, lastItemDescRow, endColumn);
                    if (tableClosingRow > 0)
                    {
                        ApplyThickBorderToRange(worksheet, tableClosingRow, startColumn, endColumn);
                        System.Diagnostics.Debug.WriteLine($"[ApplyBorders] Row {tableClosingRow}: Border applied (table closing row)");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ApplyBorders] Error: {ex.Message}");
            }
        }

        private int FindFirstSummaryRow(IXLWorksheet worksheet)
        {
            for (int row = 1; row <= 200; row++)
            {
                var cellB = worksheet.Cell(row, 2);
                var cellC = worksheet.Cell(row, 3);
                var cellD = worksheet.Cell(row, 4);

                string cellBValue = cellB.GetString().ToUpper();
                string cellCValue = cellC.GetString().ToUpper();
                string cellDValue = cellD.GetString().ToUpper();

                if (cellBValue.Contains("SUBTOTAL") || cellBValue.Contains("DISCOUNT") || cellBValue.Contains("TOTAL") ||
                    cellCValue.Contains("SUBTOTAL") || cellCValue.Contains("DISCOUNT") || cellCValue.Contains("TOTAL") ||
                    cellDValue.Contains("SUBTOTAL") || cellDValue.Contains("DISCOUNT") || cellDValue.Contains("TOTAL"))
                {
                    return row;
                }
            }
            return -1;
        }

        private int FindLastItemDescriptionRow(IXLWorksheet worksheet)
        {
            int lastItemNum = -1;
            int lastItemStartRow = -1;

            for (int row = 1; row <= 300; row++)
            {
                var cellA = worksheet.Cell(row, 1);
                string cellAValue = cellA.GetString();

                if (!string.IsNullOrEmpty(cellAValue) && int.TryParse(cellAValue.Trim(), out int itemNum))
                {
                    lastItemNum = itemNum;
                    lastItemStartRow = row;
                }
            }

            if (lastItemNum < 0) return -1;

            int lastDescRow = lastItemStartRow;
            for (int row = lastItemStartRow + 1; row <= 300; row++)
            {
                var cellA = worksheet.Cell(row, 1);
                var cellB = worksheet.Cell(row, 2);
                var cellC = worksheet.Cell(row, 3);

                string cellAValue = cellA.GetString();
                string cellBValue = cellB.GetString();
                string cellCValue = cellC.GetString();

                if (!string.IsNullOrEmpty(cellAValue) && int.TryParse(cellAValue.Trim(), out int nextItemNum) && nextItemNum > lastItemNum)
                    break;

                if (worksheet.Cell(row, 2).Style.Border.BottomBorder != XLBorderStyleValues.None)
                    break;

                if (!string.IsNullOrEmpty(cellBValue))
                {
                    lastDescRow = row;
                }
                else if (!string.IsNullOrEmpty(cellCValue) && !cellCValue.ToUpper().Contains("PC") && !cellCValue.ToUpper().Contains("UNIT"))
                {
                    lastDescRow = row;
                }
                else
                {
                    if (row > lastItemStartRow + 1 && string.IsNullOrEmpty(cellBValue) && string.IsNullOrEmpty(cellCValue))
                    {
                        bool foundMoreDesc = false;
                        for (int checkRow = row + 1; checkRow <= row + 5 && checkRow <= 300; checkRow++)
                        {
                            string checkBValue = worksheet.Cell(checkRow, 2).GetString();
                            if (!string.IsNullOrEmpty(checkBValue))
                            {
                                foundMoreDesc = true;
                                row = checkRow - 1;
                                break;
                            }
                        }

                        if (!foundMoreDesc)
                            break;
                    }
                }
            }

            return lastDescRow;
        }

        private bool IsRowInsideItemsTable(IXLWorksheet worksheet, int rowNumber, string startColumn)
        {
            try
            {
                var cell = worksheet.Cell($"{startColumn}{rowNumber}");
                var border = cell.Style.Border;

                bool hasLeftBorder = border.LeftBorder != XLBorderStyleValues.None;
                bool hasTopBorder = border.TopBorder != XLBorderStyleValues.None;
                bool hasRightBorder = border.RightBorder != XLBorderStyleValues.None;

                return hasLeftBorder || hasTopBorder || hasRightBorder;
            }
            catch
            {
                return false;
            }
        }

        private bool HasMultiplePages(IXLWorksheet worksheet)
        {
            try
            {
                var usedRange = worksheet.RangeUsed();
                if (usedRange == null) return false;

                int lastRow = usedRange.LastRow().RowNumber();
                var pageSetup = worksheet.PageSetup;
                double available = 842 - pageSetup.Margins.Top - pageSetup.Margins.Bottom;
                double total = 0;

                for (int row = 1; row <= lastRow; row++)
                {
                    double h = worksheet.Row(row).Height;
                    total += (h > 0) ? h : 15;
                }

                return total > available;
            }
            catch { return false; }
        }

        private int FindTableClosingBorderRow(IXLWorksheet worksheet, int afterRow, string endColumn)
        {
            for (int row = afterRow + 1; row <= afterRow + 10; row++)
            {
                var cellB = worksheet.Cell(row, 2);
                // Only match medium or thick borders, not thin ones from template formatting
                if (cellB.Style.Border.BottomBorder == XLBorderStyleValues.Medium ||
                    cellB.Style.Border.BottomBorder == XLBorderStyleValues.Thick)
                    return row;

                var cellA = worksheet.Cell(row, 1);
                if (cellA.Style.Border.BottomBorder == XLBorderStyleValues.Medium ||
                    cellA.Style.Border.BottomBorder == XLBorderStyleValues.Thick)
                    return row;
            }
            return afterRow + 1;
        }
        private void EnsureSubtotalTopBorder(IXLWorksheet worksheet, int subtotalRow, string startColumn, string endColumn)
        {
            if (subtotalRow <= 0) return;
            try
            {
                var range = worksheet.Range($"{startColumn}{subtotalRow}:{endColumn}{subtotalRow}");
                foreach (var cell in range.Cells())
                {
                    if (cell.Style.Border.TopBorder == XLBorderStyleValues.None)
                    {
                        cell.Style.Border.TopBorder = XLBorderStyleValues.Medium;
                        cell.Style.Border.TopBorderColor = XLColor.Black;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error applying subtotal top border: {ex.Message}");
            }
        }

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
                System.Diagnostics.Debug.WriteLine($"Error applying bottom border: {ex.Message}");
            }
        }

    }
}