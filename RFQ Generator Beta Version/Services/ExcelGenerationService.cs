using ClosedXML.Excel;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.Vml;
using RFQ_Generator_System.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RFQ_Generator_System.Services
{
    public class ExcelGenerationService
    {
        private readonly FieldMappingRepo fieldMappingRepo;
        private readonly ClientRepo clientRepo;

        public ExcelGenerationService()
        {
            fieldMappingRepo = new FieldMappingRepo();
            clientRepo = new ClientRepo();
        }

        /// <summary>
        /// Generate Excel file from template with RFQ data
        /// </summary>
        public void GenerateRFQExcel(string templatePath, string outputPath, RFQ rfq, List<RFQItem> items, int templateId)
        {
            // Load the template
            using (var workbook = new XLWorkbook(templatePath))
            {
                var worksheet = workbook.Worksheet(1); // First worksheet

                // Get field mappings for this template
                var mappings = fieldMappingRepo.GetFieldMappingsByTemplateId(templateId);
                var mappingDict = mappings.ToDictionary(m => m.FieldKey, m => m.ExcellCell);

                // Get client name
                var client = clientRepo.GetClientById(rfq.ClientId);
                string clientName = client?.ClientName ?? "";

                // Fill Header Fields
                FillHeaderFields(worksheet, mappingDict, rfq, clientName);

                // Fill Items - Each description line gets its own row
                FillItemsWithMultilineDesc(worksheet, mappingDict, items);

                // Calculate and Fill Summary using placeholder search
                FillSummaryWithPlaceholders(worksheet, items, rfq.Discount);

                // Save the file
                workbook.SaveAs(outputPath);
            }
        }

        private void FillHeaderFields(IXLWorksheet worksheet, Dictionary<string, string> mappings, RFQ rfq, string clientName)
        {
            // Client Name - "To: " + A15
            if (mappings.ContainsKey("ClientName"))
            {
                var cell = worksheet.Cell(mappings["ClientName"]);
                cell.Value = "TO : " + clientName;
            }

            // RFQ Code - "RFQ No: " + A19
            if (mappings.ContainsKey("RFQCode"))
            {
                var cell = worksheet.Cell(mappings["RFQCode"]);
                cell.Value = "RFQ NO : " + rfq.RFQCode;
            }

            // Date - ":" + G15
            if (mappings.ContainsKey("Date"))
            {
                var cell = worksheet.Cell(mappings["Date"]);
                cell.Value = ": " + rfq.CreatedAt.ToString("dd MMMM yyyy");
            }

            // Quote Code - ":" + G16
            if (mappings.ContainsKey("QuoteCode"))
            {
                var cell = worksheet.Cell(mappings["QuoteCode"]);
                cell.Value = ": " + rfq.QuoteCode;
            }

            // Delivery Term - ":" + G17
            if (mappings.ContainsKey("DeliveryTerm"))
            {
                var cell = worksheet.Cell(mappings["DeliveryTerm"]);
                cell.Value = ": " + (rfq.DeliveryTerm ?? "");
            }

            // Delivery Point - ":"
            if (mappings.ContainsKey("DeliveryPoint"))
            {
                var cell = worksheet.Cell(mappings["DeliveryPoint"]);
                cell.Value = ": " + (rfq.DeliveryPoint ?? "");
            }

            // Validity - ":" + G19
            if (mappings.ContainsKey("Validity"))
            {
                var cell = worksheet.Cell(mappings["Validity"]);
                cell.Value = ": " + (rfq.Validity ?? "");
            }
        }

        /// <summary>
        /// Each item can have multi-line descriptions, where each line takes a separate row
        /// After the last description line, there's one empty row as spacing
        /// </summary>
        private void FillItemsWithMultilineDesc(IXLWorksheet worksheet, Dictionary<string, string> mappings, List<RFQItem> items)
        {
            if (!mappings.ContainsKey("ItemStartRow"))
                return;

            // Get column mappings (with defaults if not found)
            string itemNoCol = mappings.ContainsKey("ItemNoColumn") ? mappings["ItemNoColumn"] : "A";
            string descCol = mappings.ContainsKey("ItemDescColumn") ? mappings["ItemDescColumn"] : "B";
            string deliveryCol = mappings.ContainsKey("DeliveryColumn") ? mappings["DeliveryColumn"] : "E";
            string qtyCol = mappings.ContainsKey("QuantityColumn") ? mappings["QuantityColumn"] : "F";
            string priceCol = mappings.ContainsKey("PriceColumn") ? mappings["PriceColumn"] : "G";
            string totalCol = mappings.ContainsKey("TotalColumn") ? mappings["TotalColumn"] : "H";

            int startRow = int.Parse(mappings["ItemStartRow"]);
            int rowsPerItem = mappings.ContainsKey("RowsPerItem") ? int.Parse(mappings["RowsPerItem"]) : 3;

            // Calculate total rows needed based on description lines + 1 empty row per item
            int totalRowsNeeded = 0;
            foreach (var item in items)
            {
                // Split description by newlines
                string[] descLines = (item.ItemDesc ?? "").Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                int linesCount = descLines.Length;

                // Each item needs: description lines + 1 empty row for spacing
                int rowsForThisItem = linesCount + 1;
                totalRowsNeeded += rowsForThisItem;
            }

            // Calculate how many rows to insert
            int templateRows = rowsPerItem;
            int rowsToInsert = totalRowsNeeded - templateRows;

            // Insert additional rows if needed
            if (rowsToInsert > 0)
            {
                worksheet.Row(startRow + rowsPerItem).InsertRowsAbove(rowsToInsert);

                // Copy formatting from template rows to new rows
                for (int r = 0; r < rowsToInsert; r++)
                {
                    int sourceRow = startRow + (r % rowsPerItem);
                    int targetRow = startRow + rowsPerItem + r;

                    worksheet.Row(targetRow).Height = worksheet.Row(sourceRow).Height;

                    // Copy cell styles
                    for (int col = 1; col <= 8; col++)
                    {
                        worksheet.Cell(targetRow, col).Style = worksheet.Cell(sourceRow, col).Style;
                    }
                }
            }

            // Fill each item
            int currentRow = startRow;
            foreach (var item in items)
            {
                // Split description by newlines
                string[] descLines = (item.ItemDesc ?? "").Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                int linesCount = descLines.Length;

                // Item Number - only on first row
                worksheet.Cell($"{itemNoCol}{currentRow}").Value = item.ItemNo;

                // Description lines - each on separate row
                for (int i = 0; i < descLines.Length; i++)
                {
                    worksheet.Cell($"{descCol}{currentRow + i}").Value = descLines[i];
                    worksheet.Cell($"{descCol}{currentRow + i}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                }

                // Delivery Time (in weeks) - only on first row
                int weeks = item.DeliveryTime / 7;
                worksheet.Cell($"{deliveryCol}{currentRow}").Value = weeks + " WEEKS";

                // Quantity with Unit - only on first row
                worksheet.Cell($"{qtyCol}{currentRow}").Value = item.Quantity + " " + item.UnitName;

                // Unit Price - only on first row
                worksheet.Cell($"{priceCol}{currentRow}").Value = item.UnitPrice;
                worksheet.Cell($"{priceCol}{currentRow}").Style.NumberFormat.Format = "#,##0.00";

                // Total Price (formula) - only on first row
                worksheet.Cell($"{totalCol}{currentRow}").FormulaA1 = $"={priceCol}{currentRow}*{item.Quantity}";
                worksheet.Cell($"{totalCol}{currentRow}").Style.NumberFormat.Format = "#,##0.00";

                // Move to next item row (description lines + 1 empty row for spacing)
                currentRow += linesCount + 1;
            }
        }

        private void FillSummaryWithPlaceholders(IXLWorksheet worksheet, List<RFQItem> items, decimal discount)
        {
            // Calculate subtotal (sum of all item totals)
            decimal subtotal = items.Sum(item => item.UnitPrice * item.Quantity);
            decimal totalPrice = subtotal - discount;

            // Search for placeholder cells and replace them
            var usedCells = worksheet.CellsUsed();

            foreach (var cell in usedCells)
            {
                if (cell.IsEmpty())
                    continue;

                string cellValue = cell.GetString().ToLower();

                // Check for subtotal placeholder
                if (cellValue.Contains("sub_total") || cellValue.Contains("subtotal"))
                {
                    if (cell.Address.ColumnLetter == "H")
                    {
                        cell.Value = subtotal;
                        cell.Style.NumberFormat.Format = "#,##0.00";
                    }
                }
                // Check for discount placeholder
                else if (cellValue.Contains("discount") && cellValue.Contains("_"))
                {
                    if (cell.Address.ColumnLetter == "H")
                    {
                        cell.Value = discount;
                        cell.Style.NumberFormat.Format = "#,##0.00";
                    }
                }
                // Check for total price placeholder
                else if (cellValue.Contains("total_price") || (cellValue.Contains("total") && cellValue.Contains("price") && cellValue.Contains("_")))
                {
                    if (cell.Address.ColumnLetter == "H")
                    {
                        cell.Value = totalPrice;
                        cell.Style.NumberFormat.Format = "#,##0.00";
                    }
                }
            }
        }
    }
}