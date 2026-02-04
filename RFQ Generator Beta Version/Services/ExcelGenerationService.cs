using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
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

                // Fill Item Rows with dynamic row insertion (multi-line descriptions)
                FillItemRowsWithMultiLineSupport(worksheet, mappingDict, items);

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

            //Delivery Point - ":"
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

        private void FillItemRowsWithMultiLineSupport(IXLWorksheet worksheet, Dictionary<string, string> mappings, List<RFQItem> items)
        {
            if (!mappings.ContainsKey("ItemStartRow"))
                return;

            int currentRow = int.Parse(mappings["ItemStartRow"]);
            int templateStartRow = currentRow; // Save the template start row

            // Base rows per item in template (3: desc line, unit info, blank)
            int baseRowsPerItem = 3;

            // Calculate total rows needed for ALL items
            int totalRowsNeeded = 0;
            foreach (var item in items)
            {
                // Split description by newlines
                string[] descLines = item.ItemDesc.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
                int descLineCount = descLines.Length;

                // Each item needs: description lines + 1 unit info row + 1 blank row
                int rowsForThisItem = descLineCount + 2;
                totalRowsNeeded += rowsForThisItem;
            }

            // Calculate how many rows to insert (total needed - template's 3 rows)
            int rowsToInsert = totalRowsNeeded - baseRowsPerItem;

            if (rowsToInsert > 0)
            {
                // Insert rows after the first template group
                worksheet.Row(templateStartRow + baseRowsPerItem).InsertRowsAbove(rowsToInsert);

                // Copy formatting from template rows to all new rows
                for (int r = 0; r < rowsToInsert; r++)
                {
                    int newRow = templateStartRow + baseRowsPerItem + r;

                    // Copy formatting from the first template row (row 27)
                    var templateRow = worksheet.Row(templateStartRow);
                    var newRowObj = worksheet.Row(newRow);
                    newRowObj.Height = templateRow.Height;

                    // Copy cell formatting for columns A-H
                    for (int col = 1; col <= 8; col++)
                    {
                        var templateCell = worksheet.Cell(templateStartRow, col);
                        var newCell = worksheet.Cell(newRow, col);
                        newCell.Style = templateCell.Style;
                    }
                }
            }

            // Now fill in all items with their multi-line descriptions
            currentRow = templateStartRow;

            for (int itemIndex = 0; itemIndex < items.Count; itemIndex++)
            {
                var item = items[itemIndex];

                // Split description into lines
                string[] descLines = item.ItemDesc.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

                // FIRST ROW: Item No, First line of description, Delivery Time, Quantity, Unit Price, Total Price
                worksheet.Cell($"A{currentRow}").Value = item.ItemNo;
                worksheet.Cell($"B{currentRow}").Value = descLines.Length > 0 ? descLines[0] : "";

                // Delivery Time
                worksheet.Cell($"E{currentRow}").Value = item.DeliveryTime + "WEEKS";

                // Quantity
                worksheet.Cell($"F{currentRow}").Value = item.Quantity + item.UnitName;

                // Unit Price
                worksheet.Cell($"G{currentRow}").Value = item.UnitPrice;
                worksheet.Cell($"G{currentRow}").Style.NumberFormat.Format = "#,##0.00";

                // Total Price formula
                worksheet.Cell($"H{currentRow}").FormulaA1 = $"=G{currentRow}*F{currentRow}";
                worksheet.Cell($"H{currentRow}").Style.NumberFormat.Format = "#,##0.00";

                currentRow++;

                // ADDITIONAL DESCRIPTION LINES (if any)
                for (int i = 1; i < descLines.Length; i++)
                {
                    worksheet.Cell($"B{currentRow}").Value = descLines[i];
                    currentRow++;
                }


                // BLANK/SPACING ROW (for visual separation between items)
                currentRow++;
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
    }
}