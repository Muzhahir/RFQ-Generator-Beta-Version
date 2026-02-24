using ClosedXML.Excel;
using iTextSharp.text;
using iTextSharp.text.pdf;
using RFQ_Generator_System.Repositories;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Document = iTextSharp.text.Document;
using Font = iTextSharp.text.Font;
using PageSize = iTextSharp.text.PageSize;
using Paragraph = iTextSharp.text.Paragraph;

namespace RFQ_Generator_System.Services
{
    public class PDFGenerationService
    {
        /// <summary>
        /// Generate priced PDF (default behavior)
        /// </summary>
        public string GenerateRFQPDF(string templatePath, string outputPdfPath, RFQ rfq, List<RFQItem> items, int templateId)
        {
            return GenerateRFQPDF(templatePath, outputPdfPath, rfq, items, templateId, isPriced: true);
        }

        /// <summary>
        /// Generate PDF by first creating Excel from template, then converting to PDF
        /// Uses Excel's built-in Print to PDF functionality for best results
        /// RETURNS the version suffix used (e.g., "PRICED", "UNPRICED", "COMMERCIAL", "TECHNICAL")
        /// </summary>
        /// <param name="isPriced">True for priced version, False for unpriced/technical version</param>
        public string GenerateRFQPDF(string templatePath, string outputPdfPath, RFQ rfq, List<RFQItem> items, int templateId, bool isPriced)
        {
            // Step 1: Create a temporary Excel file first
            string tempExcelPath = Path.Combine(Path.GetTempPath(), $"temp_rfq_{Guid.NewGuid()}.xlsx");

            try
            {
                // Generate Excel first using the existing service and GET the version suffix
                var excelService = new ExcelGenerationService();
                string versionSuffix = excelService.GenerateRFQExcel(templatePath, tempExcelPath, rfq, items, templateId, isPriced);

                // Step 2: Apply page break borders to the Excel file (same as Excel generation)
                ApplyPageBreakBordersToExcel(tempExcelPath, templatePath);

                // Step 3: Convert Excel to PDF - Try multiple methods
                bool success = false;

                // Method 1: Try using Microsoft.Office.Interop.Excel (best quality, preserves all formatting)
                try
                {
                    success = ConvertExcelToPDFUsingInterop(tempExcelPath, outputPdfPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Interop conversion failed: {ex.Message}");
                }

                // Method 2: If Interop failed, try using ClosedXML with improved rendering
                if (!success)
                {
                    ConvertExcelToPDFImproved(tempExcelPath, outputPdfPath);
                }

                // RETURN the version suffix
                return versionSuffix;
            }
            finally
            {
                // Clean up temporary file
                if (File.Exists(tempExcelPath))
                {
                    try
                    {
                        // Give some time for Excel to release the file
                        System.Threading.Thread.Sleep(500);
                        File.Delete(tempExcelPath);
                    }
                    catch
                    {
                        // Ignore cleanup errors
                    }
                }
            }
        }

        /// <summary>
        /// Generate BOTH priced and unpriced PDF files at the same time
        /// Returns tuple of (pricedPath, unpricedPath, pricedSuffix, unpricedSuffix)
        /// </summary>
        public (string pricedPath, string unpricedPath, string pricedSuffix, string unpricedSuffix) GenerateBothVersionsPDF(
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
            string pricedSuffix = GenerateRFQPDF(templatePath, pricedPath, rfq, items, templateId, isPriced: true);

            // Generate unpriced version
            string unpricedSuffix = GenerateRFQPDF(templatePath, unpricedPath, rfq, items, templateId, isPriced: false);

            return (pricedPath, unpricedPath, pricedSuffix, unpricedSuffix);
        }

        /// <summary>
        /// Apply page break borders to Excel file
        /// Works for multiple templates with configurable first page break row, row increment, and column range
        /// </summary>
        private void ApplyPageBreakBordersToExcel(string excelPath, string templatePath)
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
                if (!HasMultiplePages(excelPath))
                {
                    return; // Only apply borders if there are page breaks
                }

                // Open and modify the Excel file
                using (var workbook = new XLWorkbook(excelPath))
                {
                    var worksheet = workbook.Worksheet(1);

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
                        ApplyBorderToExcelRange(worksheet, pageBreakRow, startColumn, endColumn);
                        System.Diagnostics.Debug.WriteLine($"Applied border to row {pageBreakRow} ({templateFileName})");

                        // Next page break
                        pageBreakRow += rowIncrement;
                    }

                    // Save the modified workbook
                    workbook.SaveAs(excelPath);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error applying page break borders to Excel: {ex.Message}");
            }
        }

        /// <summary>
        /// Apply border to a range of cells in Excel
        /// </summary>
        private void ApplyBorderToExcelRange(IXLWorksheet worksheet, int rowNumber, string startColumn, string endColumn)
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
                System.Diagnostics.Debug.WriteLine($"Error applying border to Excel range: {ex.Message}");
            }
        }

        /// <summary>
        /// Check if the worksheet will have multiple pages when printed
        /// </summary>
        private bool HasMultiplePages(string excelPath)
        {
            try
            {
                using (var workbook = new XLWorkbook(excelPath))
                {
                    var worksheet = workbook.Worksheet(1);
                    var usedRange = worksheet.RangeUsed();
                    if (usedRange == null)
                        return false;

                    int lastRow = usedRange.LastRow().RowNumber();

                    var pageSetup = worksheet.PageSetup;
                    double topMargin = pageSetup.Margins.Top;
                    double bottomMargin = pageSetup.Margins.Bottom;

                    double pageHeight = 842;
                    double availableHeight = pageHeight - topMargin - bottomMargin;

                    double totalHeight = 0;
                    for (int row = 1; row <= lastRow; row++)
                    {
                        var rowHeight = worksheet.Row(row).Height;
                        totalHeight += (rowHeight > 0) ? rowHeight : 15;
                    }

                    return totalHeight > availableHeight;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Convert Excel to PDF using Microsoft Office Interop (BEST QUALITY - preserves exact layout)
        /// Requires Microsoft Excel to be installed on the machine
        /// </summary>
        private bool ConvertExcelToPDFUsingInterop(string excelPath, string pdfPath)
        {
            object excelApp = null;
            object workbook = null;

            try
            {
                // Check if Excel is installed by trying to create the type
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    return false; // Excel not installed
                }

                excelApp = Activator.CreateInstance(excelType);

                // Set properties using reflection to avoid dynamic issues
                excelType.InvokeMember("Visible",
                    System.Reflection.BindingFlags.SetProperty,
                    null, excelApp, new object[] { false });

                excelType.InvokeMember("DisplayAlerts",
                    System.Reflection.BindingFlags.SetProperty,
                    null, excelApp, new object[] { false });

                // Get Workbooks collection
                object workbooks = excelType.InvokeMember("Workbooks",
                    System.Reflection.BindingFlags.GetProperty,
                    null, excelApp, null);

                // Open workbook
                workbook = workbooks.GetType().InvokeMember("Open",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null, workbooks, new object[] { excelPath });

                // Export as PDF
                // XlFixedFormatType.xlTypePDF = 0
                // XlFixedFormatQuality.xlQualityStandard = 0
                workbook.GetType().InvokeMember("ExportAsFixedFormat",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null, workbook, new object[] {
                        0,      // Type: xlTypePDF
                        pdfPath, // Filename
                        0,      // Quality: xlQualityStandard
                        true,   // IncludeDocProperties
                        false,  // IgnorePrintAreas
                        Type.Missing,
                        Type.Missing,
                        false,  // OpenAfterPublish
                        Type.Missing
                    });

                // Close workbook without saving
                workbook.GetType().InvokeMember("Close",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null, workbook, new object[] { false });

                // Release workbook
                if (workbook != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workbook);
                    workbook = null;
                }

                // Release workbooks collection
                if (workbooks != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workbooks);
                    workbooks = null;
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Excel Interop error: {ex.Message}");
                return false;
            }
            finally
            {
                // Quit Excel
                if (excelApp != null)
                {
                    try
                    {
                        excelApp.GetType().InvokeMember("Quit",
                            System.Reflection.BindingFlags.InvokeMethod,
                            null, excelApp, null);
                    }
                    catch { }

                    // Release Excel application
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
                    excelApp = null;
                }

                // Force garbage collection to clean up COM objects
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        /// <summary>
        /// Improved Excel to PDF conversion that preserves more formatting
        /// This is a fallback when Excel Interop is not available
        /// </summary>
        private void ConvertExcelToPDFImproved(string excelPath, string pdfPath)
        {
            using (var workbook = new XLWorkbook(excelPath))
            {
                var worksheet = workbook.Worksheet(1);

                // Get actual page setup from Excel
                var pageSetup = worksheet.PageSetup;

                // Determine page size and orientation
                iTextSharp.text.Rectangle pageSize = PageSize.A4;
                if (pageSetup.PageOrientation == XLPageOrientation.Landscape)
                {
                    pageSize = pageSize.Rotate();
                }

                // Create PDF document with proper margins
                Document document = new Document(pageSize, 40, 40, 40, 40);

                using (FileStream fs = new FileStream(pdfPath, FileMode.Create))
                {
                    PdfWriter writer = PdfWriter.GetInstance(document, fs);
                    document.Open();

                    // Get the used range
                    var range = worksheet.RangeUsed();
                    if (range == null)
                    {
                        document.Close();
                        return;
                    }

                    int firstRow = range.FirstRow().RowNumber();
                    int lastRow = range.LastRow().RowNumber();
                    int firstCol = range.FirstColumn().ColumnNumber();
                    int lastCol = range.LastColumn().ColumnNumber();

                    int colCount = lastCol - firstCol + 1;

                    // Calculate column widths based on Excel column widths
                    float[] columnWidths = new float[colCount];
                    float totalWidth = 0;

                    for (int col = firstCol; col <= lastCol; col++)
                    {
                        float width = (float)worksheet.Column(col).Width * 7; // Approximate conversion
                        columnWidths[col - firstCol] = width;
                        totalWidth += width;
                    }

                    // Normalize widths to fit page
                    float pageWidth = pageSize.Width - document.LeftMargin - document.RightMargin;
                    for (int i = 0; i < columnWidths.Length; i++)
                    {
                        columnWidths[i] = (columnWidths[i] / totalWidth) * pageWidth;
                    }

                    // Create PDF table
                    PdfPTable table = new PdfPTable(colCount);
                    table.WidthPercentage = 100;
                    table.SetWidths(columnWidths);

                    // Process each row
                    for (int row = firstRow; row <= lastRow; row++)
                    {
                        float maxRowHeight = 0;

                        for (int col = firstCol; col <= lastCol; col++)
                        {
                            var cell = worksheet.Cell(row, col);

                            // Skip if this cell is part of a merged range but not the first cell
                            if (cell.IsMerged())
                            {
                                var mergedRange = cell.MergedRange();
                                if (mergedRange.FirstCell().Address.RowNumber != row ||
                                    mergedRange.FirstCell().Address.ColumnNumber != col)
                                {
                                    continue; // Skip merged cells that aren't the first cell
                                }
                            }

                            string cellValue = cell.GetFormattedString();

                            // Determine font
                            var xlFont = cell.Style.Font;
                            int fontSize = (int)xlFont.FontSize;
                            bool isBold = xlFont.Bold;
                            bool isItalic = xlFont.Italic;

                            int fontStyle = iTextSharp.text.Font.NORMAL;
                            if (isBold && isItalic)
                                fontStyle = iTextSharp.text.Font.BOLDITALIC;
                            else if (isBold)
                                fontStyle = iTextSharp.text.Font.BOLD;
                            else if (isItalic)
                                fontStyle = iTextSharp.text.Font.ITALIC;

                            Font cellFont = FontFactory.GetFont(FontFactory.HELVETICA, fontSize, fontStyle);

                            // Create PDF cell
                            PdfPCell pdfCell = new PdfPCell(new Phrase(cellValue, cellFont));

                            // Handle merged cells
                            if (cell.IsMerged())
                            {
                                var mergedRange = cell.MergedRange();
                                int rowspan = mergedRange.RowCount();
                                int colspan = mergedRange.ColumnCount();

                                pdfCell.Rowspan = rowspan;
                                pdfCell.Colspan = colspan;
                            }

                            // Background color
                            if (cell.Style.Fill.BackgroundColor.HasValue)
                            {
                                var bgColor = cell.Style.Fill.BackgroundColor.Color;
                                pdfCell.BackgroundColor = new BaseColor(bgColor.R, bgColor.G, bgColor.B);
                            }

                            // Alignment
                            switch (cell.Style.Alignment.Horizontal)
                            {
                                case XLAlignmentHorizontalValues.Center:
                                    pdfCell.HorizontalAlignment = Element.ALIGN_CENTER;
                                    break;
                                case XLAlignmentHorizontalValues.Right:
                                    pdfCell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                    break;
                                case XLAlignmentHorizontalValues.Left:
                                default:
                                    pdfCell.HorizontalAlignment = Element.ALIGN_LEFT;
                                    break;
                            }

                            switch (cell.Style.Alignment.Vertical)
                            {
                                case XLAlignmentVerticalValues.Center:
                                    pdfCell.VerticalAlignment = Element.ALIGN_MIDDLE;
                                    break;
                                case XLAlignmentVerticalValues.Top:
                                    pdfCell.VerticalAlignment = Element.ALIGN_TOP;
                                    break;
                                case XLAlignmentVerticalValues.Bottom:
                                    pdfCell.VerticalAlignment = Element.ALIGN_BOTTOM;
                                    break;
                            }

                            // Borders
                            if (cell.Style.Border.TopBorder != XLBorderStyleValues.None)
                                pdfCell.BorderWidthTop = 1f;
                            else
                                pdfCell.BorderWidthTop = 0f;

                            if (cell.Style.Border.BottomBorder != XLBorderStyleValues.None)
                                pdfCell.BorderWidthBottom = 1f;
                            else
                                pdfCell.BorderWidthBottom = 0f;

                            if (cell.Style.Border.LeftBorder != XLBorderStyleValues.None)
                                pdfCell.BorderWidthLeft = 1f;
                            else
                                pdfCell.BorderWidthLeft = 0f;

                            if (cell.Style.Border.RightBorder != XLBorderStyleValues.None)
                                pdfCell.BorderWidthRight = 1f;
                            else
                                pdfCell.BorderWidthRight = 0f;

                            // Padding
                            pdfCell.Padding = 5;

                            // Row height
                            float rowHeight = (float)worksheet.Row(row).Height * 1.5f;
                            if (rowHeight > maxRowHeight)
                                maxRowHeight = rowHeight;

                            table.AddCell(pdfCell);
                        }

                        // Set minimum height for the row
                        if (maxRowHeight > 0)
                        {
                            // This is handled by cell padding and content
                        }
                    }

                    document.Add(table);

                    // Add footer
                    document.Add(new Paragraph(" "));
                    Font footerFont = FontFactory.GetFont(FontFactory.HELVETICA_OBLIQUE, 8);
                    Paragraph footer = new Paragraph($"Generated on {DateTime.Now:dd/MM/yyyy HH:mm:ss}", footerFont);
                    footer.Alignment = Element.ALIGN_RIGHT;
                    document.Add(footer);

                    document.Close();
                }
            }
        }

        /// <summary>
        /// Alternative method: Generate PDF directly from RFQ data (without template)
        /// This provides more control over PDF layout
        /// </summary>
        public void GenerateDirectRFQPDF(string outputPdfPath, RFQ rfq, List<RFQItem> items, Company company, Client client)
        {
            GenerateDirectRFQPDF(outputPdfPath, rfq, items, company, client, isPriced: true);
        }

        /// <summary>
        /// Alternative method: Generate PDF directly from RFQ data (without template)
        /// This provides more control over PDF layout
        /// </summary>
        /// <param name="isPriced">True for priced version, False for unpriced/technical version</param>
        public void GenerateDirectRFQPDF(string outputPdfPath, RFQ rfq, List<RFQItem> items, Company company, Client client, bool isPriced)
        {
            Document document = new Document(PageSize.A4, 40, 40, 40, 40);

            using (FileStream fs = new FileStream(outputPdfPath, FileMode.Create))
            {
                PdfWriter writer = PdfWriter.GetInstance(document, fs);
                document.Open();

                // Fonts
                Font titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 18);
                Font headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12);
                Font labelFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10);
                Font normalFont = FontFactory.GetFont(FontFactory.HELVETICA, 10);
                Font tableHeaderFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 9);
                Font tableCellFont = FontFactory.GetFont(FontFactory.HELVETICA, 9);

                // Title
                string titleText = isPriced ? "REQUEST FOR QUOTATION" : "REQUEST FOR QUOTATION (TECHNICAL)";
                Paragraph title = new Paragraph(titleText, titleFont);
                title.Alignment = Element.ALIGN_CENTER;
                title.SpacingAfter = 20;
                document.Add(title);

                // Company Info
                PdfPTable headerTable = new PdfPTable(2);
                headerTable.WidthPercentage = 100;
                headerTable.SpacingAfter = 15;

                AddHeaderCell(headerTable, "Company:", company.CompanyName, labelFont, normalFont);
                AddHeaderCell(headerTable, "RFQ Code:", rfq.RFQCode, labelFont, normalFont);
                AddHeaderCell(headerTable, "Client:", client.ClientName, labelFont, normalFont);
                AddHeaderCell(headerTable, "Quote Code:", rfq.QuoteCode, labelFont, normalFont);
                AddHeaderCell(headerTable, "Date:", rfq.CreatedAt.ToString("dd/MM/yyyy"), labelFont, normalFont);
                AddHeaderCell(headerTable, "Validity:", rfq.Validity, labelFont, normalFont);
                AddHeaderCell(headerTable, "Delivery Point:", rfq.DeliveryPoint, labelFont, normalFont);
                AddHeaderCell(headerTable, "Delivery Terms:", rfq.DeliveryTerm, labelFont, normalFont);

                document.Add(headerTable);

                // Items Table
                Paragraph itemsTitle = new Paragraph("Items:", headerFont);
                itemsTitle.SpacingBefore = 10;
                itemsTitle.SpacingAfter = 10;
                document.Add(itemsTitle);

                PdfPTable itemsTable = new PdfPTable(6);
                itemsTable.WidthPercentage = 100;
                itemsTable.SetWidths(new float[] { 8, 35, 12, 12, 15, 18 });

                // Table headers
                AddTableHeaderCell(itemsTable, "No", tableHeaderFont);
                AddTableHeaderCell(itemsTable, "Description", tableHeaderFont);
                AddTableHeaderCell(itemsTable, "Quantity", tableHeaderFont);
                AddTableHeaderCell(itemsTable, "Unit", tableHeaderFont);

                if (isPriced)
                {
                    AddTableHeaderCell(itemsTable, "Unit Price", tableHeaderFont);
                    AddTableHeaderCell(itemsTable, "Total", tableHeaderFont);
                }
                else
                {
                    AddTableHeaderCell(itemsTable, "Unit Price", tableHeaderFont);
                    AddTableHeaderCell(itemsTable, "Total", tableHeaderFont);
                }

                // Table rows
                decimal subtotal = 0;
                foreach (var item in items)
                {
                    decimal itemTotal = item.Quantity * item.UnitPrice;
                    subtotal += itemTotal;

                    AddTableCell(itemsTable, item.ItemNo.ToString(), tableCellFont);
                    AddTableCell(itemsTable, item.ItemDesc, tableCellFont);
                    AddTableCell(itemsTable, item.Quantity.ToString(), tableCellFont);
                    AddTableCell(itemsTable, item.UnitName, tableCellFont);

                    if (isPriced)
                    {
                        AddTableCell(itemsTable, item.UnitPrice.ToString("N2"), tableCellFont);
                        AddTableCell(itemsTable, itemTotal.ToString("N2"), tableCellFont);
                    }
                    else
                    {
                        string priceText = item.UnitPrice > 0 ? "Quoted" : "Unquoted";
                        AddTableCell(itemsTable, priceText, tableCellFont);
                        AddTableCell(itemsTable, priceText, tableCellFont);
                    }
                }

                document.Add(itemsTable);

                // Totals (only for priced version)
                if (isPriced)
                {
                    PdfPTable totalsTable = new PdfPTable(2);
                    totalsTable.WidthPercentage = 40;
                    totalsTable.HorizontalAlignment = Element.ALIGN_RIGHT;
                    totalsTable.SpacingBefore = 15;

                    decimal discountAmount = subtotal * (rfq.Discount / 100);
                    decimal grandTotal = subtotal - discountAmount;

                    AddTotalRow(totalsTable, "Subtotal:", subtotal.ToString("N2"), labelFont, normalFont);
                    AddTotalRow(totalsTable, $"Discount ({rfq.Discount}%):", discountAmount.ToString("N2"), labelFont, normalFont);
                    AddTotalRow(totalsTable, "Grand Total:", grandTotal.ToString("N2"), labelFont, normalFont);

                    document.Add(totalsTable);
                }
                else
                {
                    // For unpriced version, just show "Quoted" for totals
                    PdfPTable totalsTable = new PdfPTable(2);
                    totalsTable.WidthPercentage = 40;
                    totalsTable.HorizontalAlignment = Element.ALIGN_RIGHT;
                    totalsTable.SpacingBefore = 15;

                    AddTotalRow(totalsTable, "Subtotal:", "Quoted", labelFont, normalFont);
                    AddTotalRow(totalsTable, "Discount (0%):", "0.00", labelFont, normalFont);
                    AddTotalRow(totalsTable, "Grand Total:", "Quoted", labelFont, normalFont);

                    document.Add(totalsTable);
                }

                // Footer
                Paragraph footer = new Paragraph($"\nGenerated on {DateTime.Now:dd/MM/yyyy HH:mm:ss}",
                    FontFactory.GetFont(FontFactory.HELVETICA_OBLIQUE, 8));
                footer.Alignment = Element.ALIGN_RIGHT;
                footer.SpacingBefore = 30;
                document.Add(footer);

                document.Close();
            }
        }

        private void AddHeaderCell(PdfPTable table, string label, string value, Font labelFont, Font valueFont)
        {
            PdfPCell labelCell = new PdfPCell(new Phrase(label, labelFont));
            labelCell.Border = Rectangle.NO_BORDER;
            labelCell.PaddingBottom = 5;
            table.AddCell(labelCell);

            PdfPCell valueCell = new PdfPCell(new Phrase(value, valueFont));
            valueCell.Border = Rectangle.NO_BORDER;
            valueCell.PaddingBottom = 5;
            table.AddCell(valueCell);
        }

        private void AddTableHeaderCell(PdfPTable table, string text, Font font)
        {
            PdfPCell cell = new PdfPCell(new Phrase(text, font));
            cell.BackgroundColor = new BaseColor(70, 130, 180);
            cell.HorizontalAlignment = Element.ALIGN_CENTER;
            cell.Padding = 8;
            cell.BorderColor = BaseColor.WHITE;
            table.AddCell(cell);
        }

        private void AddTableCell(PdfPTable table, string text, Font font)
        {
            PdfPCell cell = new PdfPCell(new Phrase(text, font));
            cell.Padding = 6;
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            table.AddCell(cell);
        }

        private void AddTotalRow(PdfPTable table, string label, string value, Font labelFont, Font valueFont)
        {
            PdfPCell labelCell = new PdfPCell(new Phrase(label, labelFont));
            labelCell.Border = Rectangle.NO_BORDER;
            labelCell.HorizontalAlignment = Element.ALIGN_RIGHT;
            labelCell.PaddingRight = 10;
            labelCell.PaddingTop = 5;
            labelCell.PaddingBottom = 5;
            table.AddCell(labelCell);

            PdfPCell valueCell = new PdfPCell(new Phrase(value, valueFont));
            valueCell.Border = Rectangle.NO_BORDER;
            valueCell.HorizontalAlignment = Element.ALIGN_RIGHT;
            valueCell.PaddingTop = 5;
            valueCell.PaddingBottom = 5;
            table.AddCell(valueCell);
        }
    }
}