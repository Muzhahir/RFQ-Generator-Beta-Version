

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace RFQ_Generator_System.Services
{
    public class RFQImportService
    {
        // --------------------------------------------------------
        // Data models
        // --------------------------------------------------------

        public class ImportedRFQData
        {
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
            public int DeliveryTime { get; set; } // converted from calendar days to weeks
        }

        // --------------------------------------------------------
        // Column indices (0-based) — verified against actual file
        //
        //  D (3)  = Intent to Bid              "Yes" / "No"
        //  E (4)  = Item name (short)          -> fallback desc if col L is empty
        //  G (6)  = Unit                       -> Unit  ("PC : PC; Piece" -> "PC")
        //  H (7)  = Volume                     -> Quantity
        //  K (10) = Incoterms Location         -> DeliveryPoint (primary)
        //  L (11) = Supplier Item Number       -> ItemDesc (FULL specification text)
        //  Q (16) = Supplier Inco Location     -> DeliveryPoint (fallback)
        //  R (17) = Lead Calendar Days         -> DeliveryTime (/ 7 = weeks)
        //  T (19) = Supplier Unit Price        -> UnitPrice
        //
        // Row 9  (Excel row, 1-based) = column headers
        // Row 10 (Excel row, 1-based) = first data row
        // --------------------------------------------------------

        private const int COL_INTENT = 3;   // D
        private const int COL_ITEM_NAME = 4;   // E  short name (fallback)
        private const int COL_UNIT = 6;   // G
        private const int COL_QTY = 7;   // H
        private const int COL_INCO_LOC = 10;  // K
        private const int COL_FULL_DESC = 11;  // L  full spec text (primary desc)
        private const int COL_SUP_INC_LOC = 16;  // Q
        private const int COL_LEAD_DAYS = 17;  // R
        private const int COL_UNIT_PRICE = 19;  // T

        private const int DATA_START_ROW = 10;   // Excel row number (1-based) where items begin

        // --------------------------------------------------------
        // Public API
        // --------------------------------------------------------

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

        /// <summary>
        /// Imports RFQ data directly from a .xlsx file.
        /// No NuGet packages, no COM, no Excel required.
        /// </summary>
        public ImportedRFQData ImportFromExcel(string filePath)
        {
            // Copy to temp so we can open it even if the original is locked by another process
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

        // --------------------------------------------------------
        // Core parser
        // --------------------------------------------------------

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

        // --------------------------------------------------------
        // Step 1: Load shared strings table
        // --------------------------------------------------------

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
                    // Newlines are stored as literal \r\n escape sequences in the XML text.
                    // XmlDocument reads these as characters \,r,\,n — not real line breaks.
                    // Unescape them so multiline cells come through correctly.
                    string value = sb.ToString()
                        .Replace("\\r\\n", "\r\n")
                        .Replace("\\n", "\n")
                        .Replace("\\r", "\r");
                    list.Add(value);
                }
            }
            return list;
        }

        // --------------------------------------------------------
        // Step 2: Find the data sheet file path inside the ZIP
        // --------------------------------------------------------

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

            // Map rId -> file path from workbook rels
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

            // Prefer sheet whose name starts with a digit (e.g. "1.Materials")
            // Fall back to first non-system sheet
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

        // --------------------------------------------------------
        // Step 3: Parse the worksheet XML into ImportedRFQData
        // --------------------------------------------------------

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

            bool deliveryPointRead = false;
            int itemNo = 1;

            foreach (XmlNode rowNode in doc.SelectNodes("//x:row", ns))
            {
                int rowNum = int.Parse(rowNode.Attributes["r"]?.Value ?? "0");
                if (rowNum < DATA_START_ROW) continue;

                Dictionary<int, string> cells = ReadRowCells(rowNode, ns, sharedStrings);

                // Read delivery point once from the first data row
                if (!deliveryPointRead)
                {
                    string loc = GetCell(cells, COL_INCO_LOC);
                    if (string.IsNullOrWhiteSpace(loc))
                        loc = GetCell(cells, COL_SUP_INC_LOC);
                    result.DeliveryPoint = loc?.Trim() ?? string.Empty;
                    deliveryPointRead = true;
                }

                // Only import rows marked "Yes"
                string intent = GetCell(cells, COL_INTENT);
                if (!string.Equals(intent, "Yes", StringComparison.OrdinalIgnoreCase))
                    continue;

                // Short name (col E) — used as guard and fallback
                string shortName = GetCell(cells, COL_ITEM_NAME);
                if (string.IsNullOrWhiteSpace(shortName)) continue;

                // Full specification text (col L) — primary description
                // Format: "OFFER\n\n<shortname>\nTYPE : ...\nNOMINAL SIZE : ..."
                // We strip the leading "OFFER" header line(s) to get the clean spec
                string fullDesc = GetCell(cells, COL_FULL_DESC);
                string itemDesc = BuildDescription(fullDesc, shortName);

                result.Items.Add(new ImportedItem
                {
                    ItemNo = itemNo++,
                    ItemDesc = itemDesc,
                    Unit = ParseUnit(GetCell(cells, COL_UNIT)),
                    Quantity = ParseInt(GetCell(cells, COL_QTY), 1),
                    UnitPrice = ParseDecimal(GetCell(cells, COL_UNIT_PRICE), 0m),
                    DeliveryTime = CalendarDaysToWeeks(ParseInt(GetCell(cells, COL_LEAD_DAYS), 0))
                });
            }

            return result;
        }

        // --------------------------------------------------------
        // Row cell reader
        // --------------------------------------------------------

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
                    // Shared string — rawValue is an index into the shared strings table
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
                    // Numeric or other raw value
                    value = rawValue;
                }

                dict[colIdx] = value;
            }

            return dict;
        }

        // --------------------------------------------------------
        // Description builder
        // --------------------------------------------------------

        /// <summary>
        /// Returns the full description from col L as-is (including OFFER header and all spec lines).
        /// Falls back to shortName if fullDesc is empty.
        /// </summary>
        private static string BuildDescription(string fullDesc, string shortName)
        {
            if (string.IsNullOrWhiteSpace(fullDesc))
                return shortName.Trim();

            return fullDesc.Trim();
        }

        // --------------------------------------------------------
        // Helpers
        // --------------------------------------------------------

        /// <summary>
        /// Converts a cell reference like "D10" or "AA5" to a 0-based column index.
        /// </summary>
        private static int CellRefToColIndex(string cellRef)
        {
            int result = 0;
            foreach (char ch in cellRef)
            {
                if (!char.IsLetter(ch)) break;
                result = result * 26 + (char.ToUpper(ch) - 'A' + 1);
            }
            return result - 1; // 0-based
        }

        private static string GetCell(Dictionary<int, string> cells, int colIndex)
        {
            return cells.TryGetValue(colIndex, out string val) ? val : string.Empty;
        }

        /// <summary>
        /// Parses "PC : PC; Piece" into "PC".
        /// </summary>
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