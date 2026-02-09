using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Wordprocessing;
using iTextSharp.text;
using iTextSharp.text.pdf;
using RFQ_Generator_System.Repositories;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Document = iTextSharp.text.Document;
using Font = iTextSharp.text.Font;
using PageSize = iTextSharp.text.PageSize;
using Paragraph = iTextSharp.text.Paragraph;

namespace RFQ_Generator_System.Services
{
    public class PDFGenerationService
    {
        /// <summary>
        /// Generate PDF by first creating Excel from template, then converting to PDF
        /// </summary>
        public void GenerateRFQPDF(string templatePath, string outputPdfPath, RFQ rfq, List<RFQItem> items, int templateId)
        {
            // Step 1: Create a temporary Excel file first
            string tempExcelPath = Path.Combine(Path.GetTempPath(), $"temp_rfq_{Guid.NewGuid()}.xlsx");

            try
            {
                // Generate Excel first using the existing service
                var excelService = new ExcelGenerationService();
                excelService.GenerateRFQExcel(templatePath, tempExcelPath, rfq, items, templateId);

                // Step 2: Convert Excel to PDF
                ConvertExcelToPDF(tempExcelPath, outputPdfPath);
            }
            finally
            {
                // Clean up temporary file
                if (File.Exists(tempExcelPath))
                {
                    try
                    {
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
        /// Convert Excel workbook to PDF using iTextSharp
        /// </summary>
        private void ConvertExcelToPDF(string excelPath, string pdfPath)
        {
            using (var workbook = new XLWorkbook(excelPath))
            {
                var worksheet = workbook.Worksheet(1);

                // Create PDF document
                Document document = new Document(PageSize.A4.Rotate(), 25, 25, 30, 30);

                using (FileStream fs = new FileStream(pdfPath, FileMode.Create))
                {
                    PdfWriter writer = PdfWriter.GetInstance(document, fs);
                    document.Open();

                    // Add title
                    Font titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 16);
                    Paragraph title = new Paragraph("REQUEST FOR QUOTATION", titleFont);
                    title.Alignment = Element.ALIGN_CENTER;
                    document.Add(title);
                    document.Add(new Paragraph(" ")); // Space

                    // Get the used range
                    var range = worksheet.RangeUsed();
                    int rowCount = range.RowCount();
                    int colCount = range.ColumnCount();

                    // Create PDF table
                    PdfPTable table = new PdfPTable(colCount);
                    table.WidthPercentage = 100;

                    // Define fonts
                    Font headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10);
                    Font cellFont = FontFactory.GetFont(FontFactory.HELVETICA, 9);

                    // Add cells to PDF table
                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int col = 1; col <= colCount; col++)
                        {
                            var cell = worksheet.Cell(row, col);
                            string cellValue = cell.GetString();

                            PdfPCell pdfCell = new PdfPCell(new Phrase(cellValue, row <= 2 ? headerFont : cellFont));

                            // Style header rows
                            if (row <= 2)
                            {
                                pdfCell.BackgroundColor = new BaseColor(200, 200, 200);
                                pdfCell.HorizontalAlignment = Element.ALIGN_CENTER;
                            }
                            else
                            {
                                pdfCell.HorizontalAlignment = Element.ALIGN_LEFT;
                            }

                            pdfCell.Padding = 5;
                            table.AddCell(pdfCell);
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
                Paragraph title = new Paragraph("REQUEST FOR QUOTATION", titleFont);
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
                AddTableHeaderCell(itemsTable, "Unit Price", tableHeaderFont);
                AddTableHeaderCell(itemsTable, "Total", tableHeaderFont);

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
                    AddTableCell(itemsTable, item.UnitPrice.ToString("N2"), tableCellFont);
                    AddTableCell(itemsTable, itemTotal.ToString("N2"), tableCellFont);
                }

                document.Add(itemsTable);

                // Totals
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