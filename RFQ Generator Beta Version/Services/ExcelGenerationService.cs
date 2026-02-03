using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using RFQ_Generator_System.Repositories;

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

                // Fill Item Rows with dynamic row insertion (3 rows per item)
                FillItemRowsWithDynamicInsertion(worksheet, mappingDict, items);

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

            // Validity - ":" + G19
            if (mappings.ContainsKey("Validity"))
            {
                var cell = worksheet.Cell(mappings["Validity"]);
                cell.Value = ": " + (rfq.Validity ?? "");
            }
        }

        private void FillItemRowsWithDynamicInsertion(IXLWorksheet worksheet, Dictionary<string, string> mappings, List<RFQItem> items)
        {
            if (!mappings.ContainsKey("ItemStartRow"))
                return;

            int startRow = int.Parse(mappings["ItemStartRow"]);

            // In the CG template, each item takes 3 rows (27-29 for item 1)
            // Row 1: Item No, Description, Delivery, Quantity, Unit Price, Total Price
            // Row 2: Unit/Packaging info, "WEEKS", Unit name
            // Row 3: Manufacturer/additional info
            int rowsPerItem = 3;

            // If we have more than 1 item, we need to insert additional row groups
            if (items.Count > 1)
            {
                // We need to insert (items.Count - 1) groups of 3 rows
                int rowsToInsert = (items.Count - 1) * rowsPerItem;

                // Insert rows after the first item group
                worksheet.Row(startRow + rowsPerItem).InsertRowsAbove(rowsToInsert);

                // Copy formatting from the template rows to all new rows
                for (int i = 1; i < items.Count; i++)
                {
                    int newStartRow = startRow + (i * rowsPerItem);

                    // Copy formatting from template rows (27, 28, 29) to new rows
                    for (int r = 0; r < rowsPerItem; r++)
                    {
                        var templateRow = worksheet.Row(startRow + r);
                        var newRow = worksheet.Row(newStartRow + r);
                        newRow.Height = templateRow.Height;

                        // Copy cell formatting for all columns A-H
                        for (int col = 1; col <= 8; col++)
                        {
                            var templateCell = worksheet.Cell(startRow + r, col);
                            var newCell = worksheet.Cell(newStartRow + r, col);
                            newCell.Style = templateCell.Style;
                        }
                    }
                }
            }

            // Now fill in all the items
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                int currentRow = startRow + (i * rowsPerItem);

                // Row 1 of item (e.g., row 27)
                worksheet.Cell($"A{currentRow}").Value = item.ItemNo;
                worksheet.Cell($"B{currentRow}").Value = item.ItemDesc;

                // Delivery Time - convert days to weeks
                int weeks = item.DeliveryTime / 7;
                worksheet.Cell($"E{currentRow}").Value = weeks;

                // Quantity (just the number)
                worksheet.Cell($"F{currentRow}").Value = item.Quantity;

                // Unit Price
                worksheet.Cell($"G{currentRow}").Value = item.UnitPrice;
                worksheet.Cell($"G{currentRow}").Style.NumberFormat.Format = "#,##0.00";

                // Total Price formula
                worksheet.Cell($"H{currentRow}").FormulaA1 = $"=G{currentRow}*F{currentRow}";
                worksheet.Cell($"H{currentRow}").Style.NumberFormat.Format = "#,##0.00";

                // Row 2 of item (e.g., row 28) - Unit info
                worksheet.Cell($"B{currentRow + 1}").Value = item.DeliveryTerm ?? ""; // Additional item info
                worksheet.Cell($"E{currentRow + 1}").Value = "WEEKS";
                worksheet.Cell($"F{currentRow + 1}").Value = item.UnitName;

                // Row 3 of item (e.g., row 29) - can be left for manufacturer info or notes
                // worksheet.Cell($"B{currentRow + 2}").Value = "Manufacturer :"; // Optional
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

                // Check for subtotal placeholder (looking for "sub_total", "_subtotal_", etc.)
                if (cellValue.Contains("sub_total") || cellValue.Contains("subtotal"))
                {
                    // Only replace if it's not a label (check if it's in a value column like H)
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

        // Alternative method if you still want to use field mappings for summary
        private void FillSummary(IXLWorksheet worksheet, Dictionary<string, string> mappings, List<RFQItem> items, decimal discount)
        {
            // Calculate subtotal (sum of all item totals)
            decimal subtotal = items.Sum(item => item.UnitPrice * item.Quantity);

            // Subtotal - H32
            if (mappings.ContainsKey("Subtotal"))
            {
                var cell = worksheet.Cell(mappings["Subtotal"]);
                cell.Value = subtotal;
                cell.Style.NumberFormat.Format = "#,##0.00";
            }

            // Discount - H33
            if (mappings.ContainsKey("Discount"))
            {
                var cell = worksheet.Cell(mappings["Discount"]);
                cell.Value = discount;
                cell.Style.NumberFormat.Format = "#,##0.00";
            }

            // Total Price - H34 = Subtotal - Discount
            if (mappings.ContainsKey("TotalPrice"))
            {
                var cell = worksheet.Cell(mappings["TotalPrice"]);
                decimal totalPrice = subtotal - discount;
                cell.Value = totalPrice;
                cell.Style.NumberFormat.Format = "#,##0.00";
            }
        }
    }
}