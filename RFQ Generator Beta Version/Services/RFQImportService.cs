using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;

namespace RFQ_Generator_System.Services
{
    public class RFQImportService
    {

        public class ImportedRFQData
        {
            public string RFQCode { get; set; }
            public string DeliveryPoint { get; set; }
            public string Validity { get; set; }
            public List<ImportedItem> Items { get; set; } = new List<ImportedItem>();
        }

        public class ImportedItem
        {
            public int ItemNo { get; set; }
            public string ItemDesc { get; set; }
            public string Unit { get; set; }
            public int Quantity { get; set; }
            public decimal UnitPrice { get; set; }
            public int DeliveryTime { get; set; }
        }


        private const int COL_RFQ_CODE = 2;
        private const int HEADER_ROW = 9;
        private const int DATA_START_ROW = 10;

        private class ColumnMap
        {
            public int Intent = -1;
            public int ItemName = -1;
            public int Unit = -1;
            public int Qty = -1;
            public int IncoLoc = -1;
            public int FullDesc = -1;
            public int SupIncLoc = -1;
            public int LeadDays = -1;
            public int UnitPrice = -1;
        }

        private static ColumnMap BuildColumnMap(Dictionary<int, string> headerRow)
        {
            var m = new ColumnMap();
            foreach (var kvp in headerRow)
            {
                string h = kvp.Value.Trim();
                int col = kvp.Key;

                if (h.Equals("*Intent to Bid", StringComparison.OrdinalIgnoreCase)) m.Intent = col;
                else if (h.Equals("*Item name", StringComparison.OrdinalIgnoreCase)) m.ItemName = col;
                else if (h.Equals("*Unit", StringComparison.OrdinalIgnoreCase)) m.Unit = col;
                else if (h.Equals("*Volume", StringComparison.OrdinalIgnoreCase)) m.Qty = col;
                else if (h.Equals("Incoterms Location", StringComparison.OrdinalIgnoreCase)) m.IncoLoc = col;
                else if (h.Equals("Supplier Item Number", StringComparison.OrdinalIgnoreCase)) m.FullDesc = col;
                else if (h.Equals("*Supplier Incoterms Location", StringComparison.OrdinalIgnoreCase)) m.SupIncLoc = col;
                else if (h.Equals("*Estimated Lead Calendar Days", StringComparison.OrdinalIgnoreCase)) m.LeadDays = col;
                else if (h.StartsWith("*Supplier Unit Price", StringComparison.OrdinalIgnoreCase)) m.UnitPrice = col;
            }
            return m;
        }


        public static string ShowFileDialog()
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Select Excel Pricesheet to Import";
                ofd.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                ofd.DefaultExt = ".xlsx";
                return ofd.ShowDialog() == DialogResult.OK ? ofd.FileName : null;
            }
        }

        public bool IsValidExcelFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return false;

            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            return ext == ".xlsx";
        }

        public ImportedRFQData ImportFromExcel(string filePath)
        {
            string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            try
            {
                File.Copy(filePath, tempPath, overwrite: true);
                return ParseXlsx(tempPath);
            }
            finally
            {
                try { File.Delete(tempPath); } catch { }
            }
        }



        private ImportedRFQData ParseXlsx(string filePath)
        {
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                List<string> sharedStrings = LoadSharedStrings(zip);
                string sheetFile = FindDataSheetFile(zip);

                if (sheetFile == null)
                    throw new Exception(
                        "Could not find a data sheet in this file.\n" +
                        "Expected a sheet like '1.Materials'.");

                return ParseSheet(zip, sheetFile, sharedStrings);
            }
        }



        private List<string> LoadSharedStrings(ZipArchive zip)
        {
            var list = new List<string>();
            ZipArchiveEntry entry = zip.GetEntry("xl/sharedStrings.xml");
            if (entry == null) return list;

            using (Stream s = entry.Open())
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(s);

                XmlNamespaceManager ns = new XmlNamespaceManager(doc.NameTable);
                ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                foreach (XmlNode si in doc.SelectNodes("//x:si", ns))
                {
                    StringBuilder sb = new StringBuilder();
                    foreach (XmlNode t in si.SelectNodes(".//x:t", ns))
                    {
                        if (t.InnerText != null)
                            sb.Append(t.InnerText);
                    }

                    string value = sb.ToString()
                        .Replace("\\r\\n", "\r\n")
                        .Replace("\\n", "\n")
                        .Replace("\\r", "\r");
                    list.Add(value);
                }
            }
            return list;
        }



        private string FindDataSheetFile(ZipArchive zip)
        {
            ZipArchiveEntry wbEntry = zip.GetEntry("xl/workbook.xml");
            if (wbEntry == null) return null;

            XmlDocument wbDoc = new XmlDocument();
            using (Stream s = wbEntry.Open())
                wbDoc.Load(s);

            XmlNamespaceManager wbNs = new XmlNamespaceManager(wbDoc.NameTable);
            wbNs.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            wbNs.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            Dictionary<string, string> rIdToFile = new Dictionary<string, string>();
            ZipArchiveEntry relsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
            if (relsEntry != null)
            {
                XmlDocument relsDoc = new XmlDocument();
                using (Stream s = relsEntry.Open())
                    relsDoc.Load(s);

                foreach (XmlNode rel in relsDoc.SelectNodes("//*[local-name()='Relationship']"))
                {
                    string id = rel.Attributes["Id"]?.Value;
                    string target = rel.Attributes["Target"]?.Value;
                    if (id != null && target != null)
                        rIdToFile[id] = "xl/" + target.TrimStart('/');
                }
            }


            string fallbackRId = null;
            HashSet<string> systemNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                { "Instructions", "Sheet", "teehsecirP" };

            foreach (XmlNode sheet in wbDoc.SelectNodes("//x:sheet", wbNs))
            {
                string name = sheet.Attributes["name"]?.Value ?? "";
                string rId = sheet.Attributes["r:id"]?.Value ?? "";

                if (name.Length > 0 && char.IsDigit(name[0]))
                {
                    if (rIdToFile.TryGetValue(rId, out string file))
                        return file;
                }

                if (!systemNames.Contains(name) && fallbackRId == null)
                    fallbackRId = rId;
            }

            if (fallbackRId != null && rIdToFile.TryGetValue(fallbackRId, out string fallback))
                return fallback;

            return null;
        }



        private ImportedRFQData ParseSheet(ZipArchive zip, string sheetFile, List<string> sharedStrings)
        {
            var result = new ImportedRFQData();

            ZipArchiveEntry entry = zip.GetEntry(sheetFile);
            if (entry == null)
                throw new Exception($"Sheet file not found in xlsx: {sheetFile}");

            XmlDocument doc = new XmlDocument();
            using (Stream s = entry.Open())
                doc.Load(s);

            XmlNamespaceManager ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            ColumnMap cols = null;
            bool deliveryPointRead = false;
            int itemNo = 1;

            foreach (XmlNode rowNode in doc.SelectNodes("//x:row", ns))
            {
                int rowNum = int.Parse(rowNode.Attributes["r"]?.Value ?? "0");

                if (rowNum == 1)
                {
                    var r1 = ReadRowCells(rowNode, ns, sharedStrings);
                    result.RFQCode = ExtractRFQCode(GetCell(r1, COL_RFQ_CODE));
                    continue;
                }

                if (rowNum == HEADER_ROW)
                {
                    cols = BuildColumnMap(ReadRowCells(rowNode, ns, sharedStrings));
                    continue;
                }

                if (rowNum < DATA_START_ROW) continue;
                if (cols == null) continue;

                Dictionary<int, string> cells = ReadRowCells(rowNode, ns, sharedStrings);

                if (!deliveryPointRead)
                {
                    string loc = cols.IncoLoc >= 0 ? GetCell(cells, cols.IncoLoc) : string.Empty;
                    if (string.IsNullOrWhiteSpace(loc) && cols.SupIncLoc >= 0)
                        loc = GetCell(cells, cols.SupIncLoc);
                    result.DeliveryPoint = loc?.Trim() ?? string.Empty;
                    deliveryPointRead = true;
                }

                string intent = cols.Intent >= 0 ? GetCell(cells, cols.Intent) : string.Empty;
                if (!string.Equals(intent, "Yes", StringComparison.OrdinalIgnoreCase))
                    continue;

                string shortName = cols.ItemName >= 0 ? GetCell(cells, cols.ItemName) : string.Empty;
                if (string.IsNullOrWhiteSpace(shortName)) continue;

                string fullDesc = cols.FullDesc >= 0 ? GetCell(cells, cols.FullDesc) : string.Empty;
                string itemDesc = BuildDescription(fullDesc, shortName);

                result.Items.Add(new ImportedItem
                {
                    ItemNo = itemNo++,
                    ItemDesc = itemDesc,
                    Unit = ParseUnit(cols.Unit >= 0 ? GetCell(cells, cols.Unit) : string.Empty),
                    Quantity = ParseInt(cols.Qty >= 0 ? GetCell(cells, cols.Qty) : string.Empty, 1),
                    UnitPrice = ParseDecimal(cols.UnitPrice >= 0 ? GetCell(cells, cols.UnitPrice) : string.Empty, 0m),
                    DeliveryTime = CalendarDaysToWeeks(ParseInt(cols.LeadDays >= 0 ? GetCell(cells, cols.LeadDays) : string.Empty, 0))
                });
            }

            return result;
        }


        private Dictionary<int, string> ReadRowCells(XmlNode rowNode, XmlNamespaceManager ns,
            List<string> sharedStrings)
        {
            var dict = new Dictionary<int, string>();

            foreach (XmlNode c in rowNode.SelectNodes("x:c", ns))
            {
                string cellRef = c.Attributes["r"]?.Value ?? "";
                int colIdx = CellRefToColIndex(cellRef);
                string cellType = c.Attributes["t"]?.Value ?? "";

                XmlNode vNode = c.SelectSingleNode("x:v", ns);
                string rawValue = vNode?.InnerText;

                if (rawValue == null) continue;

                string value;
                if (cellType == "s")
                {
                    if (int.TryParse(rawValue, out int idx) && idx < sharedStrings.Count)
                        value = sharedStrings[idx];
                    else
                        value = string.Empty;
                }
                else if (cellType == "inlineStr")
                {
                    XmlNode isNode = c.SelectSingleNode("x:is/x:t", ns);
                    value = isNode?.InnerText ?? string.Empty;
                }
                else
                {
                    value = rawValue;
                }

                dict[colIdx] = value;
            }

            return dict;
        }




        private static string ExtractRFQCode(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;

            int sepIdx = raw.IndexOf(" - ", StringComparison.Ordinal);
            string code = sepIdx > 0 ? raw.Substring(0, sepIdx) : raw;
            return code.Trim();
        }


        // Builds the item description from the full description field, falling back to
        // the short name if the full description is empty.
        //
        // FIX: Some PETRONAS pricesheet cells contain the same spec block repeated
        // multiple times. This can occur as either:
        //   (a) blank-line-separated paragraphs  (\r\n\r\n / \n\n)
        //   (b) single-newline-separated lines   (\r\n / \n)  ← root cause for Item 1 border bug
        //
        // Both cases are now deduplicated so that no item description is inflated
        // beyond its true line count, which was causing a spurious page-break bottom
        // border to appear mid-item in the generated quotation.
        private static string BuildDescription(string fullDesc, string shortName)
        {
            if (string.IsNullOrWhiteSpace(fullDesc))
                return shortName.Trim();

            string trimmed = fullDesc.Trim();

            // ----------------------------------------------------------------
            // Pass 1: try to deduplicate at the paragraph (blank-line) level.
            // ----------------------------------------------------------------
            string[] paragraphs = trimmed.Split(
                new[] { "\r\n\r\n", "\n\n", "\r\r" },
                StringSplitOptions.RemoveEmptyEntries);

            if (paragraphs.Length > 1)
            {
                var seenParas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var uniqueParas = new List<string>();
                foreach (string para in paragraphs)
                {
                    string key = Regex.Replace(para.Trim(), @"\s+", " ");
                    if (seenParas.Add(key))
                        uniqueParas.Add(para.Trim());
                }

                // If deduplication removed anything, return the cleaned result.
                if (uniqueParas.Count < paragraphs.Length)
                    return uniqueParas.Count == 1
                        ? uniqueParas[0]
                        : string.Join("\r\n", uniqueParas);

                // No duplicates at paragraph level — return as-is (joined with single newline).
                return string.Join("\r\n", uniqueParas);
            }

            // ----------------------------------------------------------------
            // Pass 2: no blank-line separators found.
            // Try deduplication at the individual line level.
            // This catches the case where the same spec lines are repeated
            // with only \r\n between them (the root cause of the Item 1 bug).
            // ----------------------------------------------------------------
            string[] lines = trimmed.Split(
                new[] { "\r\n", "\r", "\n" },
                StringSplitOptions.RemoveEmptyEntries);

            if (lines.Length > 1)
            {
                var seenLines = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var uniqueLines = new List<string>();
                foreach (string line in lines)
                {
                    string key = Regex.Replace(line.Trim(), @"\s+", " ");
                    if (seenLines.Add(key))
                        uniqueLines.Add(line.Trim());
                }

                // Only use the deduplicated result if lines were actually removed.
                // If every line was unique, preserve the original (no false positives).
                if (uniqueLines.Count < lines.Length)
                    return string.Join("\r\n", uniqueLines);
            }

            // Nothing to deduplicate — return the trimmed original.
            return trimmed;
        }



        private static int CellRefToColIndex(string cellRef)
        {
            int result = 0;
            foreach (char ch in cellRef)
            {
                if (!char.IsLetter(ch)) break;
                result = result * 26 + (char.ToUpper(ch) - 'A' + 1);
            }
            return result - 1;
        }

        private static string GetCell(Dictionary<int, string> cells, int colIndex)
        {
            return cells.TryGetValue(colIndex, out string val) ? val : string.Empty;
        }

        private static string ParseUnit(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "PC";
            int idx = raw.IndexOf(" : ", StringComparison.Ordinal);
            if (idx > 0) return raw.Substring(0, idx).Trim();
            string first = raw.Split(' ')[0].Trim();
            return string.IsNullOrEmpty(first) ? "PC" : first;
        }

        private static int CalendarDaysToWeeks(int days)
        {
            return days <= 0 ? 0 : (int)Math.Round(days / 7.0);
        }

        private static int ParseInt(string s, int defaultValue = 0)
        {
            if (string.IsNullOrWhiteSpace(s)) return defaultValue;
            if (decimal.TryParse(s, out decimal d)) return (int)d;
            return defaultValue;
        }

        private static decimal ParseDecimal(string s, decimal defaultValue = 0m)
        {
            if (string.IsNullOrWhiteSpace(s)) return defaultValue;
            if (decimal.TryParse(s, out decimal d)) return d;
            return defaultValue;
        }
    }
}