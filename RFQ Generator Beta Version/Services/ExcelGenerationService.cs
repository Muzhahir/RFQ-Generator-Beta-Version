using ClosedXML.Excel;
using RFQ_Generator_System.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RFQ_Generator_System.Services
{
    public class ExcelGenerationService
    {
        private readonly ClientRepo clientRepo;
        private readonly RFQItemRepo rfqItemRepo; // <- Added RFQItemRepo

        // HARDCODED PLACEHOLDERS - No database needed!
        private static class Placeholders
        {
            // Header placeholders
            public const string ClientName = "client_name";
            public const string Date = "_date";
            public const string RFQCode = "rfq_code";
            public const string QuoteCode = "quote_code";
            public const string Validity = "_validity";
            public const string PaymentTerm = "payment_term";
            public const string DeliveryTerm = "delivery_term";
            public const string DeliveryPoint = "delivery_point";

            // Item table column placeholders
            public const string ItemNo = "_num";
            public const string ItemDesc = "_desc";
            public const string DeliveryTime = "delivery_time";
            public const string Quantity = "_qty";
            public const string UnitPrice = "unit_price";
            public const string TotalPrice = "total_price";

            // Summary placeholders
            public const string Subtotal = "sub_total";
            public const string Discount = "_discount";
            public const string SummaryTotal = "summary_total_price";
        }

        public ExcelGenerationService()
        {
            clientRepo = new ClientRepo();
            rfqItemRepo = new RFQItemRepo(); // <- Initialize RFQItemRepo
        }

        /// <summary>
        /// Generate Excel file from template with RFQ data
        /// Templates just need to have the standard placeholders in their cells
        /// </summary>
        public void GenerateRFQExcel(string templatePath, string outputPath, RFQ rfq, List<RFQItem> items, int templateId)
        {
            using (var workbook = new XLWorkbook(templatePath))
            {
                var worksheet = workbook.Worksheet(1);

                // Get client name
                var client = clientRepo.GetClientById(rfq.ClientId);
                string clientName = client?.ClientName ?? "";

                // Find item table start row
                var itemStartCell = FindPlaceholderCell(worksheet, Placeholders.ItemNo);

                // Fill all sections
                FillHeaderFields(worksheet, rfq, clientName);

                if (itemStartCell != null)
                {
                    FillItems(worksheet, items, itemStartCell);
                }

                FillSummary(worksheet, items, rfq.Discount);

                // Save
                workbook.SaveAs(outputPath);
            }
        }

        /// <summary>
        /// Find a cell containing the specified placeholder text
        /// </summary>
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

        /// <summary>
        /// Fill header fields by finding and replacing placeholders
        /// </summary>
        private void FillHeaderFields(IXLWorksheet worksheet, RFQ rfq, string clientName)
        {
            var usedCells = worksheet.CellsUsed();

            foreach (var cell in usedCells)
            {
                if (cell.IsEmpty())
                    continue;

                string cellValue = cell.GetString();

                // Match placeholder and fill with data
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
                
            }
        }

        /// <summary>
        /// Fill items by finding placeholder columns and filling data
        /// </summary>
        private void FillItems(IXLWorksheet worksheet, List<RFQItem> items, IXLCell startCell)
        {
            int headerRow = startCell.Address.RowNumber;

            // Find columns by searching for placeholders in header row
            string itemNoCol = FindColumnByPlaceholder(worksheet, headerRow, Placeholders.ItemNo);
            string descCol = FindColumnByPlaceholder(worksheet, headerRow, Placeholders.ItemDesc);
            string deliveryCol = FindColumnByPlaceholder(worksheet, headerRow, Placeholders.DeliveryTime);
            string qtyCol = FindColumnByPlaceholder(worksheet, headerRow, Placeholders.Quantity);
            string priceCol = FindColumnByPlaceholder(worksheet, headerRow, Placeholders.UnitPrice);
            string totalCol = FindColumnByPlaceholder(worksheet, headerRow, Placeholders.TotalPrice);

            int startRow = headerRow;

            // Calculate total rows needed (multi-line descriptions)
            int totalRowsNeeded = 0;
            foreach (var item in items)
            {
                string[] descLines = (item.ItemDesc ?? "").Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                totalRowsNeeded += descLines.Length + 1; // +1 for spacing
            }

            // Insert rows if needed
            int templateRows = 1;
            int rowsToInsert = totalRowsNeeded - templateRows;

            if (rowsToInsert > 0)
            {
                worksheet.Row(startRow + templateRows).InsertRowsAbove(rowsToInsert);

                // Copy formatting
                for (int r = 0; r < rowsToInsert; r++)
                {
                    int targetRow = startRow + templateRows + r;
                    worksheet.Row(targetRow).Height = worksheet.Row(startRow).Height;

                    foreach (var cell in worksheet.Row(startRow).Cells())
                    {
                        if (!cell.IsEmpty() || cell.Style.Fill.BackgroundColor.HasValue)
                        {
                            worksheet.Cell(targetRow, cell.Address.ColumnNumber).Style = cell.Style;
                        }
                    }
                }
            }

            // Fill each item
            int currentRow = startRow;
            foreach (var item in items)
            {
                string[] descLines = (item.ItemDesc ?? "").Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

                // Item Number
                if (itemNoCol != null)
                {
                    worksheet.Cell($"{itemNoCol}{currentRow}").Value = item.ItemNo;
                }

                // Description (multi-line)
                if (descCol != null)
                {
                    for (int i = 0; i < descLines.Length; i++)
                    {
                        worksheet.Cell($"{descCol}{currentRow + i}").Value = descLines[i];
                        worksheet.Cell($"{descCol}{currentRow + i}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    }
                }

                // Delivery Time
                if (deliveryCol != null)
                {
                    int weeks = item.DeliveryTime / 7;
                    worksheet.Cell($"{deliveryCol}{currentRow}").Value = weeks + " WEEKS";
                }

                // Quantity
                if (qtyCol != null)
                {
                    worksheet.Cell($"{qtyCol}{currentRow}").Value = item.Quantity + " " + item.UnitName;
                }

                // Unit Price
                if (priceCol != null)
                {
                    worksheet.Cell($"{priceCol}{currentRow}").Value = item.UnitPrice;
                    worksheet.Cell($"{priceCol}{currentRow}").Style.NumberFormat.Format = "#,##0.00";
                }

                // Total Price (formula)
                if (totalCol != null && priceCol != null)
                {
                    worksheet.Cell($"{totalCol}{currentRow}").FormulaA1 = $"={priceCol}{currentRow}*{item.Quantity}";
                    worksheet.Cell($"{totalCol}{currentRow}").Style.NumberFormat.Format = "#,##0.00";
                }

                currentRow += descLines.Length + 1;
            }
        }

        /// <summary>
        /// Find column by searching for placeholder in specific row
        /// </summary>
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

        /// <summary>
        /// Fill summary by finding placeholders and replacing with calculations
        /// </summary>
        private void FillSummary(IXLWorksheet worksheet, List<RFQItem> items, decimal discount)
        {
            decimal subtotal = items.Sum(item => item.UnitPrice * item.Quantity);
            decimal totalPrice = subtotal - discount;

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
                    cell.Value = discount;
                    cell.Style.NumberFormat.Format = "#,##0.00";
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