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
        private static readonly HashSet<string> SameDeliveryHideTemplates = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "GA TEMPLATE.xlsx",
            "MA TEMPLATE.xlsx",
            "SC TEMPLATE.xlsx",
        };
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
                bool templateHasDeliveryPoint = HasPlaceholder(worksheet, Placeholders.DeliveryPoint);
                System.Diagnostics.Debug.WriteLine($"[ExcelService] Template terminology: {pricedTerm}/{unpricedTerm}");
                System.Diagnostics.Debug.WriteLine($"[ExcelService] Version suffix: {versionSuffix}");
                System.Diagnostics.Debug.WriteLine($"[ExcelService] Currency: {effectiveCurrency}");
                System.Diagnostics.Debug.WriteLine($"[ExcelService] Is Priced: {isPriced}");
                System.Diagnostics.Debug.WriteLine($"[ExcelService] Template has delivery_point: {templateHasDeliveryPoint}");
                if (!isPriced)
                    ConvertToUnpricedVersion(worksheet, pricedTerm, unpricedTerm);
                string templateFileName = Path.GetFileName(fullTemplatePath);
                FillHeaderFields(worksheet, rfq, clientName, items, templateHasDeliveryPoint, templateFileName);
                bool suppressDeliveryColumn = ShouldSuppressDeliveryColumn(templateFileName, items);
                if (itemStartCell != null)
                    FillItems(worksheet, items, itemStartCell, effectiveCurrency, isPriced, suppressDeliveryColumn);
                decimal effectiveDiscount = isPriced ? rfq.Discount : 0;
                var (subtotalRow, summaryTotalRow) = FillSummary(worksheet, items, effectiveDiscount, isPriced);
                var (insertedFrom, insertedTo) = EnsureTAndCFitsOnPage(worksheet, fullTemplatePath, subtotalRow, summaryTotalRow);
                ReplaceCurrencyInWorksheet(worksheet, effectiveCurrency);
                ApplyBottomBordersAtPageBreaks(worksheet, fullTemplatePath, insertedFrom, insertedTo, subtotalRow);
                workbook.SaveAs(outputPath);
                return versionSuffix;
            }
        }
        private bool ShouldSuppressDeliveryColumn(string templateFileName, List<RFQItem> items)
        {
            if (!SameDeliveryHideTemplates.Contains(templateFileName)) return false;
            if (items == null || items.Count == 0) return false;
            int first = items[0].DeliveryTime;
            return items.All(i => i.DeliveryTime == first);
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
        private bool HasPlaceholder(IXLWorksheet worksheet, string placeholder)
        {
            foreach (var cell in worksheet.CellsUsed())
                if (cell.GetString().Equals(placeholder, StringComparison.OrdinalIgnoreCase))
                    return true;
            return false;
        }
        private void FillHeaderFields(IXLWorksheet worksheet, RFQ rfq, string clientName, List<RFQItem> items, bool templateHasDeliveryPoint, string templateFileName)
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
            bool templateHasDeliveryTerm = HasPlaceholder(worksheet, Placeholders.DeliveryTerm);
            string effectiveDeliveryTerm = rfq.DeliveryTerm ?? "";
            if (!templateHasDeliveryPoint && !string.IsNullOrWhiteSpace(rfq.DeliveryPoint))
                effectiveDeliveryTerm = $"{effectiveDeliveryTerm} {rfq.DeliveryPoint}".Trim();
            string effectiveDeliveryPoint = rfq.DeliveryPoint ?? "";
            if (!templateHasDeliveryTerm && !string.IsNullOrWhiteSpace(rfq.DeliveryTerm))
                effectiveDeliveryPoint = string.IsNullOrWhiteSpace(effectiveDeliveryPoint)
                    ? rfq.DeliveryTerm
                    : $"{rfq.DeliveryTerm} {effectiveDeliveryPoint}";
            foreach (var cell in worksheet.CellsUsed())
            {
                if (cell.IsEmpty()) continue;
                string cellValue = cell.GetString();
                if (cellValue.Equals(Placeholders.ClientName, StringComparison.OrdinalIgnoreCase)) cell.Value = clientName;
                else if (cellValue.Equals(Placeholders.RFQCode, StringComparison.OrdinalIgnoreCase)) cell.Value = rfq.RFQCode;
                else if (cellValue.Equals(Placeholders.Date, StringComparison.OrdinalIgnoreCase))
                {
                    string dateFormat = templateFileName.Equals("GA TEMPLATE.xlsx", StringComparison.OrdinalIgnoreCase)
                        ? "dddd, d MMMM yyyy"
                        : "dd MMMM yyyy";
                    cell.Value = rfq.CreatedAt.ToString(dateFormat);
                }
                else if (cellValue.Equals(Placeholders.QuoteCode, StringComparison.OrdinalIgnoreCase)) cell.Value = rfq.QuoteCode;
                else if (cellValue.Equals(Placeholders.DeliveryTerm, StringComparison.OrdinalIgnoreCase)) cell.Value = effectiveDeliveryTerm;
                else if (cellValue.Equals(Placeholders.DeliveryPoint, StringComparison.OrdinalIgnoreCase)) cell.Value = effectiveDeliveryPoint;
                else if (cellValue.Equals(Placeholders.Validity, StringComparison.OrdinalIgnoreCase)) cell.Value = rfq.Validity ?? "";
                else if (cellValue.Equals(Placeholders.HeaderDeliveryTime, StringComparison.OrdinalIgnoreCase)) cell.Value = headerDeliveryTimeText;
            }
        }
        private void FillItems(IXLWorksheet worksheet, List<RFQItem> items, IXLCell startCell, string effectiveCurrency, bool isPriced, bool suppressDeliveryColumn)
        {
            int headerRow = startCell.Address.RowNumber;
            string itemNoCol = FindColumnByPlaceholder(worksheet, headerRow, Placeholders.ItemNo);
            string descCol = FindColumnByPlaceholder(worksheet, headerRow, Placeholders.ItemDesc);
            string deliveryCol = FindColumnByPlaceholder(worksheet, headerRow, Placeholders.DeliveryTime);
            string qtyCol = FindColumnByPlaceholder(worksheet, headerRow, Placeholders.Quantity);
            string priceCol = FindColumnByPlaceholder(worksheet, headerRow, Placeholders.UnitPrice);
            string totalCol = FindColumnByPlaceholder(worksheet, headerRow, Placeholders.TotalPrice);
            bool hasDeliveryColumn = deliveryCol != null && !suppressDeliveryColumn;
            int startRow = headerRow;
            int totalRowsNeeded = 0;
            foreach (var item in items)
            {
                string[] descLines = (item.ItemDesc ?? "").Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                totalRowsNeeded += descLines.Length + 1;
                if (!hasDeliveryColumn && !suppressDeliveryColumn) totalRowsNeeded += 1;
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
                    if (!hasDeliveryColumn && !suppressDeliveryColumn)
                    {
                        int deliveryTextRow = currentRow + descLines.Length;
                        worksheet.Cell($"{descCol}{deliveryTextRow}").Value = $"(Delivery: {weeks} {weekText})";
                        worksheet.Cell($"{descCol}{deliveryTextRow}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                        worksheet.Cell($"{descCol}{deliveryTextRow}").Style.Font.Italic = true;
                    }
                }
                if (hasDeliveryColumn)
                {
                    worksheet.Cell($"{deliveryCol}{currentRow}").Value = weeks;
                    worksheet.Cell($"{deliveryCol}{currentRow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell($"{deliveryCol}{currentRow}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                    worksheet.Cell($"{deliveryCol}{currentRow + 1}").Value = weekText;
                    worksheet.Cell($"{deliveryCol}{currentRow + 1}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell($"{deliveryCol}{currentRow + 1}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                }
                if (qtyCol != null)
                {
                    worksheet.Cell($"{qtyCol}{currentRow}").Value = item.Quantity;
                    worksheet.Cell($"{qtyCol}{currentRow}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell($"{qtyCol}{currentRow}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Bottom;
                    worksheet.Cell($"{qtyCol}{currentRow + 1}").Value = item.UnitName;
                    worksheet.Cell($"{qtyCol}{currentRow + 1}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell($"{qtyCol}{currentRow + 1}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                }
                if (priceCol != null)
                {
                    if (isPriced)
                    {
                        worksheet.Cell($"{priceCol}{currentRow}").Value = item.UnitPrice;
                        worksheet.Cell($"{priceCol}{currentRow}").Style.NumberFormat.Format = "#,##0.00";
                    }
                    else
                    {
                        worksheet.Cell($"{priceCol}{currentRow}").Value = item.UnitPrice > 0 ? "QUOTED" : "UNQUOTED";
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
                        worksheet.Cell($"{totalCol}{currentRow}").Value = item.UnitPrice > 0 ? "QUOTED" : "UNQUOTED";
                        worksheet.Cell($"{totalCol}{currentRow}").Style.NumberFormat.Format = "@";
                    }
                }
                int rowAdvance = descLines.Length + 1;
                if (!hasDeliveryColumn && !suppressDeliveryColumn) rowAdvance += 1;
                currentRow += rowAdvance;
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
                    else { cell.Value = "QUOTED"; cell.Style.NumberFormat.Format = "@"; }
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
                    else { cell.Value = "QUOTED"; cell.Style.NumberFormat.Format = "@"; }
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
                    { "GA TEMPLATE.xlsx",   23 },
                    { "MA TEMPLATE.xlsx",   26 },
                    { "OGIT TEMPLATE.xlsx", 29 },
                    { "OP TEMPLATE.xlsx",   33 },
                    { "PO TEMPLATE.xlsx",   25 },
                    { "SC TEMPLATE.xlsx",   24 },
                };
                var templateConfig = new Dictionary<string, (int firstRow, int increment)>(StringComparer.OrdinalIgnoreCase)
                {
                    { "CG TEMPLATE.xlsx",   (68, 43) },
                    { "DE TEMPLATE.xlsx",   (60, 43) },
                    { "GA TEMPLATE.xlsx",   (65, 42) },
                    { "MA TEMPLATE.xlsx",   (61, 44) },
                    { "OGIT TEMPLATE.xlsx", (67, 46) },
                    { "OP TEMPLATE.xlsx",   (70, 50) },
                    { "PO TEMPLATE.xlsx",   (75, 73) },
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
        private void ApplyBottomBordersAtPageBreaks(IXLWorksheet worksheet, string templatePath, int insertedFrom, int insertedTo, int subtotalRow)
        {
            try
            {
                string templateFileName = Path.GetFileName(templatePath);
                var templateConfig = new Dictionary<string, (int firstRow, int increment, string startCol, string endCol)>(StringComparer.OrdinalIgnoreCase)
                {
                    { "CG TEMPLATE.xlsx",   (68, 43, "A", "J") },
                    { "DE TEMPLATE.xlsx",   (60, 43, "A", "T") },
                    { "GA TEMPLATE.xlsx",   (65, 42, "A", "H") },
                    { "MA TEMPLATE.xlsx",   (61, 44, "A", "H") },
                    { "OGIT TEMPLATE.xlsx", (67, 46, "A", "H") },
                    { "OP TEMPLATE.xlsx",   (70, 50, "A", "H") },
                    { "PO TEMPLATE.xlsx",   (75, 73, "A", "I") },
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
                int firstSummaryRow = FindFirstSummaryRow(worksheet);
                System.Diagnostics.Debug.WriteLine($"[ApplyBorders] First summary row: {firstSummaryRow}");
                int itemsTableEndRow = subtotalRow > 0 ? subtotalRow - 1 : (firstSummaryRow > 0 ? firstSummaryRow - 1 : -1);
                System.Diagnostics.Debug.WriteLine($"[ApplyBorders] Items table end row: {itemsTableEndRow}");
                EnsureSubtotalTopBorder(worksheet, firstSummaryRow, startColumn, endColumn);
                var pageBreakRows = new HashSet<int>();
                int pb = firstPageBreakRow;
                while (pb <= lastRow)
                {
                    pageBreakRows.Add(pb);
                    pb += rowIncrement;
                }
                if (!HasMultiplePages(worksheet))
                {
                    System.Diagnostics.Debug.WriteLine($"[ApplyBorders] Single page, no page break borders needed");
                    return;
                }
                System.Diagnostics.Debug.WriteLine($"[ApplyBorders] Multi-page document, applying borders only at page breaks inside items table");
                foreach (int pageBreakRow in pageBreakRows.OrderBy(r => r))
                {
                    if (insertedFrom >= 0 && pageBreakRow >= insertedFrom && pageBreakRow <= insertedTo)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ApplyBorders] Row {pageBreakRow}: Skipped (inside inserted blank rows)");
                        continue;
                    }
                    if (itemsTableEndRow > 0 && pageBreakRow > itemsTableEndRow)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ApplyBorders] Row {pageBreakRow}: Skipped (past items table end row {itemsTableEndRow})");
                        continue;
                    }
                    if (!IsRowInsideItemsTable(worksheet, pageBreakRow, startColumn))
                    {
                        System.Diagnostics.Debug.WriteLine($"[ApplyBorders] Row {pageBreakRow}: Skipped (not in items table)");
                        continue;
                    }
                    ApplyThickBorderToRange(worksheet, pageBreakRow, startColumn, endColumn);
                    System.Diagnostics.Debug.WriteLine($"[ApplyBorders] Row {pageBreakRow}: Border applied");
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
                    cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
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