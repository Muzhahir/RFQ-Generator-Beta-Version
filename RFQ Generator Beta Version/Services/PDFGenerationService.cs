using ClosedXML.Excel;
using iTextSharp.text;
using iTextSharp.text.pdf;
using RFQ_Generator_System.Repositories;
using System;
using System.Collections.Generic;
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


        public string GenerateRFQPDF(string templatePath, string outputPdfPath, RFQ rfq, List<RFQItem> items, int templateId)
        {
            return GenerateRFQPDF(templatePath, outputPdfPath, rfq, items, templateId, isPriced: true);
        }


        public string GenerateRFQPDF(string templatePath, string outputPdfPath, RFQ rfq, List<RFQItem> items, int templateId, bool isPriced)
        {
            string tempExcelPath = Path.Combine(Path.GetTempPath(), $"temp-rfq-{Guid.NewGuid()}.xlsx");

            try
            {
                var excelService = new ExcelGenerationService();
                string versionSuffix = excelService.GenerateRFQExcel(templatePath, tempExcelPath, rfq, items, templateId, isPriced);

                bool success = false;
                try
                {
                    success = ConvertExcelToPDFUsingInterop(tempExcelPath, outputPdfPath);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Interop conversion failed: {ex.Message}");
                }

                if (!success)
                    ConvertExcelToPDFImproved(tempExcelPath, outputPdfPath);

                return versionSuffix;
            }
            finally
            {
                try
                {
                    System.Threading.Thread.Sleep(500);
                    if (File.Exists(tempExcelPath))
                        File.Delete(tempExcelPath);
                }
                catch { }
            }
        }


        public (string pricedPath, string unpricedPath, string pricedSuffix, string unpricedSuffix) GenerateBothVersionsPDF(
            string templatePath,
            string baseOutputPath,
            RFQ rfq,
            List<RFQItem> items,
            int templateId)
        {
            string directory = Path.GetDirectoryName(baseOutputPath);
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(baseOutputPath);
            string extension = Path.GetExtension(baseOutputPath);

            string tempPricedPath = Path.Combine(directory, $"{fileNameWithoutExt}-TEMP-PRICED{extension}");
            string tempUnpricedPath = Path.Combine(directory, $"{fileNameWithoutExt}-TEMP-UNPRICED{extension}");

            string pricedSuffix = GenerateRFQPDF(templatePath, tempPricedPath, rfq, items, templateId, isPriced: true);
            string unpricedSuffix = GenerateRFQPDF(templatePath, tempUnpricedPath, rfq, items, templateId, isPriced: false);

            string pricedPath = Path.Combine(directory, $"{fileNameWithoutExt}-{pricedSuffix}{extension}");
            string unpricedPath = Path.Combine(directory, $"{fileNameWithoutExt}-{unpricedSuffix}{extension}");

            if (File.Exists(pricedPath)) File.Delete(pricedPath);
            if (File.Exists(unpricedPath)) File.Delete(unpricedPath);

            File.Move(tempPricedPath, pricedPath);
            File.Move(tempUnpricedPath, unpricedPath);

            return (pricedPath, unpricedPath, pricedSuffix, unpricedSuffix);
        }



        public void GenerateDirectRFQPDF(string outputPdfPath, RFQ rfq, List<RFQItem> items, Company company, Client client)
        {
            GenerateDirectRFQPDF(outputPdfPath, rfq, items, company, client, isPriced: true);
        }


        public void GenerateDirectRFQPDF(string outputPdfPath, RFQ rfq, List<RFQItem> items, Company company, Client client, bool isPriced)
        {

            string pricedTerm = "PRICED";
            string unpricedTerm = "UNPRICED";
            string versionLabel = isPriced ? pricedTerm : unpricedTerm;


            string currency = string.IsNullOrWhiteSpace(rfq.Currency) ? "RM" : rfq.Currency;


            var deliveryWeeks = items
                .Select(i => i.DeliveryTime)
                .Distinct()
                .OrderBy(w => w)
                .ToList();

            string headerDeliveryTimeText = "";
            if (deliveryWeeks.Count > 0)
            {
                int maxWeeks = deliveryWeeks.Max();
                string wLabel = maxWeeks < 2 ? "WEEK" : "WEEKS";
                headerDeliveryTimeText = string.Join(", ", deliveryWeeks) + " " + wLabel;
            }

            decimal effectiveDiscount = isPriced ? rfq.Discount : 0m;
            decimal subtotal = items.Sum(i => i.UnitPrice * i.Quantity);
            decimal discountAmount = subtotal * effectiveDiscount / 100m;
            decimal grandTotal = subtotal - discountAmount;

            Document document = new Document(PageSize.A4, 40, 40, 40, 40);

            using (FileStream fs = new FileStream(outputPdfPath, FileMode.Create))
            {
                PdfWriter.GetInstance(document, fs);
                document.Open();

                Font titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 18);
                Font headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12);
                Font labelFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10);
                Font normalFont = FontFactory.GetFont(FontFactory.HELVETICA, 10);
                Font tableHeaderFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 9);
                Font tableCellFont = FontFactory.GetFont(FontFactory.HELVETICA, 9);
                Font italicFont = FontFactory.GetFont(FontFactory.HELVETICA_OBLIQUE, 9);

                string titleText = $"REQUEST FOR QUOTATION ({versionLabel})";
                Paragraph title = new Paragraph(titleText, titleFont)
                {
                    Alignment = Element.ALIGN_CENTER,
                    SpacingAfter = 20
                };
                document.Add(title);


                PdfPTable headerTable = new PdfPTable(2)
                {
                    WidthPercentage = 100,
                    SpacingAfter = 15
                };

                AddHeaderCell(headerTable, "Company:", company.CompanyName, labelFont, normalFont);
                AddHeaderCell(headerTable, "RFQ Code:", rfq.RFQCode, labelFont, normalFont);
                AddHeaderCell(headerTable, "Client:", client.ClientName, labelFont, normalFont);
                AddHeaderCell(headerTable, "Quote Code:", rfq.QuoteCode, labelFont, normalFont);
                AddHeaderCell(headerTable, "Date:", rfq.CreatedAt.ToString("dd MMMM yyyy"), labelFont, normalFont);
                AddHeaderCell(headerTable, "Validity:", rfq.Validity ?? "", labelFont, normalFont);
                AddHeaderCell(headerTable, "Delivery Point:", rfq.DeliveryPoint ?? "", labelFont, normalFont);
                AddHeaderCell(headerTable, "Delivery Terms:", rfq.DeliveryTerm ?? "", labelFont, normalFont);
                AddHeaderCell(headerTable, "Delivery Time:", headerDeliveryTimeText, labelFont, normalFont);  

                document.Add(headerTable);

              
                Paragraph itemsTitle = new Paragraph("Items:", headerFont)
                {
                    SpacingBefore = 10,
                    SpacingAfter = 10
                };
                document.Add(itemsTitle);

                
                PdfPTable itemsTable = new PdfPTable(6)
                {
                    WidthPercentage = 100
                };
                itemsTable.SetWidths(new float[] { 6, 32, 12, 14, 16, 20 });

                AddTableHeaderCell(itemsTable, "No", tableHeaderFont);
                AddTableHeaderCell(itemsTable, "Description", tableHeaderFont);
                AddTableHeaderCell(itemsTable, "Delivery", tableHeaderFont);
                AddTableHeaderCell(itemsTable, "Qty", tableHeaderFont);
                AddTableHeaderCell(itemsTable, $"Unit Price ({currency})", tableHeaderFont);  
                AddTableHeaderCell(itemsTable, $"Total ({currency})", tableHeaderFont);  

                foreach (var item in items)
                {
                    
                    string[] descLines = (item.ItemDesc ?? "").Split(
                        new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

                    int weeks = item.DeliveryTime;
                    string weekText = weeks < 2 ? "WEEK" : "WEEKS";
                    string deliveryStr = $"{weeks} {weekText}";

                    string fullDesc = string.Join("\n", descLines);

                    PdfPCell noCell = new PdfPCell(new Phrase(item.ItemNo.ToString(), tableCellFont))
                    {
                        Padding = 6,
                        HorizontalAlignment = Element.ALIGN_CENTER
                    };
                    itemsTable.AddCell(noCell);

                    PdfPCell descCell = new PdfPCell(new Phrase(fullDesc, tableCellFont))
                    {
                        Padding = 6,
                        HorizontalAlignment = Element.ALIGN_LEFT
                    };
                    itemsTable.AddCell(descCell);

                    AddTableCell(itemsTable, deliveryStr, tableCellFont);

                    AddTableCell(itemsTable, $"{item.Quantity} {item.UnitName}", tableCellFont);

                    if (isPriced)
                    {
                        AddTableCell(itemsTable, item.UnitPrice.ToString("N2"), tableCellFont);
                        AddTableCell(itemsTable, (item.Quantity * item.UnitPrice).ToString("N2"), tableCellFont);
                    }
                    else
                    {
                        string tag = item.UnitPrice > 0 ? "QUOTED" : "UNQUOTED";
                        AddTableCell(itemsTable, tag, tableCellFont);
                        AddTableCell(itemsTable, tag, tableCellFont);
                    }
                }

                document.Add(itemsTable);

                PdfPTable totalsTable = new PdfPTable(2)
                {
                    WidthPercentage = 45,
                    HorizontalAlignment = Element.ALIGN_RIGHT,
                    SpacingBefore = 15
                };

                if (isPriced)
                {
                    AddTotalRow(totalsTable, "Subtotal:", subtotal.ToString("N2"), labelFont, normalFont);
                    AddTotalRow(totalsTable, $"Discount ({effectiveDiscount:0.##}%):", discountAmount.ToString("N2"), labelFont, normalFont);
                    AddTotalRow(totalsTable, $"Grand Total ({currency}):", grandTotal.ToString("N2"), labelFont, normalFont);
                }
                else
                {
                    AddTotalRow(totalsTable, "Subtotal:", "QUOTED", labelFont, normalFont);
                    AddTotalRow(totalsTable, "Discount (0%):", "0.00", labelFont, normalFont);
                    AddTotalRow(totalsTable, $"Grand Total ({currency}):", "QUOTED", labelFont, normalFont);
                }

                document.Add(totalsTable);

                Paragraph footer = new Paragraph(
                    $"\nGenerated on {DateTime.Now:dd/MM/yyyy HH:mm:ss}",
                    FontFactory.GetFont(FontFactory.HELVETICA_OBLIQUE, 8))
                {
                    Alignment = Element.ALIGN_RIGHT,
                    SpacingBefore = 30
                };
                document.Add(footer);

                document.Close();
            }
        }


        private bool ConvertExcelToPDFUsingInterop(string excelPath, string pdfPath)
        {
            object excelApp = null;
            object workbook = null;
            object workbooks = null;

            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) return false;

                excelApp = Activator.CreateInstance(excelType);

                excelType.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, excelApp, new object[] { false });
                excelType.InvokeMember("DisplayAlerts", System.Reflection.BindingFlags.SetProperty, null, excelApp, new object[] { false });

                workbooks = excelType.InvokeMember("Workbooks", System.Reflection.BindingFlags.GetProperty, null, excelApp, null);

                workbook = workbooks.GetType().InvokeMember("Open",
                    System.Reflection.BindingFlags.InvokeMethod, null, workbooks,
                    new object[] { excelPath });

                workbook.GetType().InvokeMember("ExportAsFixedFormat",
                    System.Reflection.BindingFlags.InvokeMethod, null, workbook,
                    new object[] { 0, pdfPath, 0, true, false,
                                   Type.Missing, Type.Missing, false, Type.Missing });

                workbook.GetType().InvokeMember("Close",
                    System.Reflection.BindingFlags.InvokeMethod, null, workbook,
                    new object[] { false });

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Excel Interop error: {ex.Message}");
                return false;
            }
            finally
            {
                try
                {
                    if (workbook != null)
                    {
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workbook);
                        workbook = null;
                    }
                    if (workbooks != null)
                    {
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workbooks);
                        workbooks = null;
                    }
                    if (excelApp != null)
                    {
                        excelApp.GetType().InvokeMember("Quit",
                            System.Reflection.BindingFlags.InvokeMethod, null, excelApp, null);
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
                        excelApp = null;
                    }
                }
                catch { }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private void ConvertExcelToPDFImproved(string excelPath, string pdfPath)
        {
            using (var workbook = new XLWorkbook(excelPath))
            {
                var worksheet = workbook.Worksheet(1);
                var pageSetup = worksheet.PageSetup;

                iTextSharp.text.Rectangle pageSize = PageSize.A4;
                if (pageSetup.PageOrientation == XLPageOrientation.Landscape)
                    pageSize = pageSize.Rotate();

                Document document = new Document(pageSize, 40, 40, 40, 40);

                using (FileStream fs = new FileStream(pdfPath, FileMode.Create))
                {
                    PdfWriter.GetInstance(document, fs);
                    document.Open();

                    var range = worksheet.RangeUsed();
                    if (range == null) { document.Close(); return; }

                    int firstRow = range.FirstRow().RowNumber();
                    int lastRow = range.LastRow().RowNumber();
                    int firstCol = range.FirstColumn().ColumnNumber();
                    int lastCol = range.LastColumn().ColumnNumber();
                    int colCount = lastCol - firstCol + 1;

                    float[] columnWidths = new float[colCount];
                    float totalWidth = 0;

                    for (int col = firstCol; col <= lastCol; col++)
                    {
                        float w = (float)worksheet.Column(col).Width * 7;
                        columnWidths[col - firstCol] = w;
                        totalWidth += w;
                    }

                    float pageWidth = pageSize.Width - document.LeftMargin - document.RightMargin;
                    for (int i = 0; i < columnWidths.Length; i++)
                        columnWidths[i] = (columnWidths[i] / totalWidth) * pageWidth;

                    PdfPTable table = new PdfPTable(colCount)
                    {
                        WidthPercentage = 100
                    };
                    table.SetWidths(columnWidths);

                    for (int row = firstRow; row <= lastRow; row++)
                    {
                        for (int col = firstCol; col <= lastCol; col++)
                        {
                            var cell = worksheet.Cell(row, col);

                            if (cell.IsMerged())
                            {
                                var mr = cell.MergedRange();
                                if (mr.FirstCell().Address.RowNumber != row ||
                                    mr.FirstCell().Address.ColumnNumber != col)
                                    continue;
                            }

                            string cellValue = cell.GetFormattedString();
                            var xlFont = cell.Style.Font;
                            int fSize = (int)xlFont.FontSize;
                            bool bold = xlFont.Bold;
                            bool italic = xlFont.Italic;

                            int fontStyle = iTextSharp.text.Font.NORMAL;
                            if (bold && italic) fontStyle = iTextSharp.text.Font.BOLDITALIC;
                            else if (bold) fontStyle = iTextSharp.text.Font.BOLD;
                            else if (italic) fontStyle = iTextSharp.text.Font.ITALIC;

                            Font cellFont = FontFactory.GetFont(FontFactory.HELVETICA, fSize, fontStyle);
                            PdfPCell pdfCell = new PdfPCell(new Phrase(cellValue, cellFont));

                            if (cell.IsMerged())
                            {
                                var mr = cell.MergedRange();
                                pdfCell.Rowspan = mr.RowCount();
                                pdfCell.Colspan = mr.ColumnCount();
                            }

                            if (cell.Style.Fill.BackgroundColor.HasValue)
                            {
                                var bg = cell.Style.Fill.BackgroundColor.Color;
                                pdfCell.BackgroundColor = new BaseColor(bg.R, bg.G, bg.B);
                            }

                            switch (cell.Style.Alignment.Horizontal)
                            {
                                case XLAlignmentHorizontalValues.Center: pdfCell.HorizontalAlignment = Element.ALIGN_CENTER; break;
                                case XLAlignmentHorizontalValues.Right: pdfCell.HorizontalAlignment = Element.ALIGN_RIGHT; break;
                                default: pdfCell.HorizontalAlignment = Element.ALIGN_LEFT; break;
                            }

                            switch (cell.Style.Alignment.Vertical)
                            {
                                case XLAlignmentVerticalValues.Center: pdfCell.VerticalAlignment = Element.ALIGN_MIDDLE; break;
                                case XLAlignmentVerticalValues.Top: pdfCell.VerticalAlignment = Element.ALIGN_TOP; break;
                                default: pdfCell.VerticalAlignment = Element.ALIGN_BOTTOM; break;
                            }

                            pdfCell.BorderWidthTop = cell.Style.Border.TopBorder != XLBorderStyleValues.None ? 1f : 0f;
                            pdfCell.BorderWidthBottom = cell.Style.Border.BottomBorder != XLBorderStyleValues.None ? 1f : 0f;
                            pdfCell.BorderWidthLeft = cell.Style.Border.LeftBorder != XLBorderStyleValues.None ? 1f : 0f;
                            pdfCell.BorderWidthRight = cell.Style.Border.RightBorder != XLBorderStyleValues.None ? 1f : 0f;

                            pdfCell.Padding = 5;
                            table.AddCell(pdfCell);
                        }
                    }

                    document.Add(table);

                    Paragraph footer = new Paragraph(
                        $"Generated on {DateTime.Now:dd/MM/yyyy HH:mm:ss}",
                        FontFactory.GetFont(FontFactory.HELVETICA_OBLIQUE, 8))
                    {
                        Alignment = Element.ALIGN_RIGHT
                    };
                    document.Add(new Paragraph(" "));
                    document.Add(footer);

                    document.Close();
                }
            }
        }


        private void AddHeaderCell(PdfPTable table, string label, string value, Font labelFont, Font valueFont)
        {
            PdfPCell labelCell = new PdfPCell(new Phrase(label, labelFont))
            {
                Border = Rectangle.NO_BORDER,
                PaddingBottom = 5
            };
            table.AddCell(labelCell);

            PdfPCell valueCell = new PdfPCell(new Phrase(value ?? "", valueFont))
            {
                Border = Rectangle.NO_BORDER,
                PaddingBottom = 5
            };
            table.AddCell(valueCell);
        }

        private void AddTableHeaderCell(PdfPTable table, string text, Font font)
        {
            PdfPCell cell = new PdfPCell(new Phrase(text, font))
            {
                BackgroundColor = new BaseColor(70, 130, 180),
                HorizontalAlignment = Element.ALIGN_CENTER,
                Padding = 8,
                BorderColor = BaseColor.WHITE
            };
            table.AddCell(cell);
        }

        private void AddTableCell(PdfPTable table, string text, Font font)
        {
            PdfPCell cell = new PdfPCell(new Phrase(text ?? "", font))
            {
                Padding = 6,
                HorizontalAlignment = Element.ALIGN_LEFT
            };
            table.AddCell(cell);
        }

        private void AddTotalRow(PdfPTable table, string label, string value, Font labelFont, Font valueFont)
        {
            PdfPCell labelCell = new PdfPCell(new Phrase(label, labelFont))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = Element.ALIGN_RIGHT,
                PaddingRight = 10,
                PaddingTop = 5,
                PaddingBottom = 5
            };
            table.AddCell(labelCell);

            PdfPCell valueCell = new PdfPCell(new Phrase(value ?? "", valueFont))
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = Element.ALIGN_RIGHT,
                PaddingTop = 5,
                PaddingBottom = 5
            };
            table.AddCell(valueCell);
        }
    }
}