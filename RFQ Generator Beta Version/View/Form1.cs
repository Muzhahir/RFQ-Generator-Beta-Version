using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using RFQ_Generator_System.Services;
using RFQ_Generator_System.Repositories;

namespace RFQ_Generator_System
{
    public partial class Form1 : Form
    {
        private RFQService rfqService;
        private ExcelGenerationService excelService;
        private PDFGenerationService pdfService;
        private RFQImportService importService;
        private Template currentTemplate;
        private int currentRFQId = 0;
        private List<RFQItem> rfqItems;
        private int editingItemIndex = -1;
        private List<Company> allCompanies;
        private List<Client> allClients;
        private bool isQuoteCodeManuallyEdited = false;
        private bool isRFQSaved = false;

        // ✅ Stores the folder of the last imported pricesheet.
        // When set, Excel/PDF save dialogs will default to this folder instead of the network folder.
        private string importedFileFolder = null;

        // ✅ Network save folder - change this if the path ever changes
        private const string RFQOutputFolder = @"\\DLINK-01731D\Volume_1\Public_J\8. Operation Dept\RFQ";

        public Form1()
        {
            InitializeComponent();
            rfqService = new RFQService();
            excelService = new ExcelGenerationService();
            pdfService = new PDFGenerationService();
            importService = new RFQImportService();
            rfqItems = new List<RFQItem>();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                cmbCompany.DropDownStyle = ComboBoxStyle.DropDown;
                cmbClient.DropDownStyle = ComboBoxStyle.DropDown;

                LoadCompanies();
                LoadClients();

                dtpCreatedAt.Value = DateTime.Now;
                cmbCurrency.SelectedIndex = 0;

                rbtnDAP.Checked = true;

                ClearItemForm();
                UpdateItemsList();

                txtQuoteCode.ReadOnly = false;
                txtQuoteCode.BackColor = Color.LightYellow;

                cmbCompany.SelectedIndexChanged += HandleCompanyOrClientChange;
                cmbClient.SelectedIndexChanged += HandleCompanyOrClientChange;
                txtQuoteCode.TextChanged += txtQuoteCode_TextChanged;

                // ✅ Enable drag and drop
                this.AllowDrop = true;
                this.DragEnter += Form1_DragEnter;
                this.DragDrop += Form1_DragDrop;
            }
            catch (Exception ex)
            {
                MessageBox.Show("The app failed to start properly. Please restart and try again.",
                    "Startup Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region Data Loading
        private void LoadCompanies()
        {
            try
            {
                allCompanies = rfqService.GetAllCompanies();
                var companiesWithPlaceholder = new List<Company>
                {
                    new Company { Id = 0, CompanyName = "-- Select Company --", CompanyCode = "" }
                };
                companiesWithPlaceholder.AddRange(allCompanies);
                cmbCompany.DataSource = companiesWithPlaceholder;
                cmbCompany.DisplayMember = "CompanyName";
                cmbCompany.ValueMember = "Id";
                cmbCompany.SelectedIndex = 0;
                cmbCompany.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                cmbCompany.AutoCompleteSource = AutoCompleteSource.ListItems;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not load the company list. Please check your database connection and restart.",
                    "Load Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadClients()
        {
            try
            {
                allClients = rfqService.GetAllClients();
                var sortedClients = allClients.OrderBy(c => c.ClientName).ToList();

                var clientsWithPlaceholder = new List<Client>
                {
                    new Client { Id = 0, ClientName = "-- Select Client --", ClientCode = "" }
                };
                clientsWithPlaceholder.AddRange(sortedClients);

                cmbClient.DataSource = clientsWithPlaceholder;
                cmbClient.DisplayMember = "DisplayText";
                cmbClient.ValueMember = "Id";
                cmbClient.DropDownStyle = ComboBoxStyle.DropDown;
                cmbClient.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                cmbClient.AutoCompleteSource = AutoCompleteSource.ListItems;
                cmbClient.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not load the client list. Please check your database connection and restart.",
                    "Load Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region Quote Code Generation
        private void txtQuoteCode_TextChanged(object sender, EventArgs e)
        {
            if (txtQuoteCode.Focused)
            {
                isQuoteCodeManuallyEdited = true;
                txtQuoteCode.BackColor = Color.LightGreen;
            }
        }

        private void HandleCompanyOrClientChange(object sender, EventArgs e)
        {
            if (sender == cmbCompany)
            {
                UpdateTemplateDisplay();
                rfqService.ClearQuoteCodeCache();
                isQuoteCodeManuallyEdited = false;
                isRFQSaved = false;

                // ✅ Reset RFQ ID when company changes so a new record is created
                currentRFQId = 0;
            }

            if (sender == cmbClient)
            {
                rfqService.ClearQuoteCodeCache();
                isQuoteCodeManuallyEdited = false;
                isRFQSaved = false;

                // ✅ Reset RFQ ID when client changes so a new record is created
                currentRFQId = 0;

                // ✅ Auto-set delivery term based on selected client
                if (cmbClient.SelectedValue is int clientId && clientId > 0)
                {
                    var selectedClient = allClients.FirstOrDefault(c => c.Id == clientId);
                    if (selectedClient != null && !string.IsNullOrEmpty(selectedClient.DeliveryTerm))
                    {
                        switch (selectedClient.DeliveryTerm.Trim().ToUpper())
                        {
                            case "DAP": rbtnDAP.Checked = true; break;
                            case "DDP": rbtnDDP.Checked = true; break;
                            case "DDU/DAP": rbtnDDUDAP.Checked = true; break;
                            case "PCG": rbtnPCG.Checked = true; break;
                            default: rbtnDAP.Checked = true; break;
                        }
                    }
                }
            }

            if (!txtQuoteCode.Focused)
                GenerateAndDisplayQuoteCodePreview();
        }

        private void UpdateTemplateDisplay()
        {
            if (cmbCompany.SelectedValue == null || !(cmbCompany.SelectedValue is int) || (int)cmbCompany.SelectedValue == 0)
            {
                lblTemplatePath.Text = "No template selected";
                lblTemplatePath.ForeColor = Color.Gray;
                currentTemplate = null;
                return;
            }

            try
            {
                int companyId = (int)cmbCompany.SelectedValue;
                currentTemplate = rfqService.GetTemplateByCompanyId(companyId);
                if (currentTemplate != null)
                {
                    lblTemplatePath.Text = currentTemplate.TemplatePath;
                    lblTemplatePath.ForeColor = Color.Green;
                }
                else
                {
                    lblTemplatePath.Text = "No template found for this company";
                    lblTemplatePath.ForeColor = Color.Red;
                    currentTemplate = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not load the template for this company. Please try selecting the company again.",
                    "Template Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GenerateAndDisplayQuoteCodePreview()
        {
            try
            {
                if (cmbCompany.SelectedValue == null || !(cmbCompany.SelectedValue is int) || (int)cmbCompany.SelectedValue == 0)
                {
                    txtQuoteCode.Text = "";
                    txtQuoteCode.BackColor = Color.LightYellow;
                    return;
                }

                if (cmbClient.SelectedValue == null || !(cmbClient.SelectedValue is int) || (int)cmbClient.SelectedValue == 0)
                {
                    txtQuoteCode.Text = "";
                    txtQuoteCode.BackColor = Color.LightYellow;
                    return;
                }

                int companyId = (int)cmbCompany.SelectedValue;
                int clientId = (int)cmbClient.SelectedValue;

                var selectedCompany = allCompanies.FirstOrDefault(c => c.Id == companyId);
                var selectedClient = allClients.FirstOrDefault(c => c.Id == clientId);

                if (selectedCompany != null && selectedClient != null)
                {
                    string companyCode = selectedCompany.CompanyCode;
                    string clientCode = selectedClient.ClientCode ?? "";

                    if (string.IsNullOrEmpty(companyCode))
                    {
                        txtQuoteCode.Text = "[Company code missing]";
                        txtQuoteCode.BackColor = Color.LightCoral;
                        MessageBox.Show(
                            $"'{selectedCompany.CompanyName}' is missing a company code in the database.\n\nPlease contact your administrator to set it up.",
                            "Missing Company Code", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    string quoteCode = rfqService.GenerateQuoteCodePreview(companyCode, clientCode);
                    if (string.IsNullOrEmpty(quoteCode))
                    {
                        txtQuoteCode.Text = "[Could not generate code]";
                        txtQuoteCode.BackColor = Color.LightCoral;
                    }
                    else
                    {
                        txtQuoteCode.Text = quoteCode;
                        txtQuoteCode.BackColor = Color.LightYellow;
                    }
                }
            }
            catch (Exception ex)
            {
                txtQuoteCode.Text = "[Error]";
                txtQuoteCode.BackColor = Color.LightCoral;
                MessageBox.Show("Could not generate the quote code. Please reselect the company and client and try again.",
                    "Quote Code Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        private void cmbCompany_SelectedIndexChanged(object sender, EventArgs e)
        {
            HandleCompanyOrClientChange(sender, e);
        }

        #region Item Management
        private void btnAddItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ValidateItemForm()) return;

                var item = new RFQItem
                {
                    ItemDesc = txtItemDescription.Text.Trim(),
                    Quantity = (int)numItemQuantity.Value,
                    UnitName = txtItemUnit.Text.Trim(),
                    UnitPrice = numItemUnitPrice.Value,
                    DeliveryTime = (int)numItemDeliveryTime.Value
                };

                if (editingItemIndex >= 0)
                {
                    rfqItems[editingItemIndex] = item;
                    editingItemIndex = -1;
                    btnAddItem.Text = "Add Item";
                }
                else
                {
                    rfqItems.Add(item);
                }

                for (int i = 0; i < rfqItems.Count; i++)
                    rfqItems[i].ItemNo = i + 1;

                UpdateItemsList();
                ClearItemForm();
                txtItemDescription.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not add the item. Please check the details and try again.",
                    "Add Item Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnClearItem_Click(object sender, EventArgs e)
        {
            ClearItemForm();
            editingItemIndex = -1;
            btnAddItem.Text = "Add Item";
        }

        private void btnEditItem_Click(object sender, EventArgs e)
        {
            if (lstItems.SelectedIndex < 0)
            {
                MessageBox.Show("Please select an item from the list to edit.",
                    "No Item Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            editingItemIndex = lstItems.SelectedIndex;
            var item = rfqItems[editingItemIndex];
            txtItemDescription.Text = item.ItemDesc;
            numItemQuantity.Value = item.Quantity;
            txtItemUnit.Text = item.UnitName;
            numItemUnitPrice.Value = item.UnitPrice;
            numItemDeliveryTime.Value = item.DeliveryTime;
            btnAddItem.Text = "Update Item";
            txtItemDescription.Focus();
        }

        private void btnRemoveItem_Click(object sender, EventArgs e)
        {
            if (lstItems.SelectedIndex < 0)
            {
                MessageBox.Show("Please select an item from the list to remove.",
                    "No Item Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var result = MessageBox.Show("Remove this item? This cannot be undone.",
                "Remove Item", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                rfqItems.RemoveAt(lstItems.SelectedIndex);
                for (int i = 0; i < rfqItems.Count; i++)
                    rfqItems[i].ItemNo = i + 1;
                UpdateItemsList();
                ClearItemForm();
            }
        }

        private void lstItems_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool hasSelection = lstItems.SelectedIndex >= 0;
            btnEditItem.Enabled = hasSelection;
            btnRemoveItem.Enabled = hasSelection;
        }

        private void UpdateItemsList()
        {
            lstItems.Items.Clear();
            foreach (var item in rfqItems)
            {
                string shortDesc = item.ItemDesc.Replace("\r\n", " ").Replace("\n", " ");
                if (shortDesc.Length > 40) shortDesc = shortDesc.Substring(0, 40) + "...";
                string itemText = $"{item.ItemNo}. {shortDesc} | Qty: {item.Quantity} {item.UnitName} | " +
                                  $"Price: {item.UnitPrice:N2} | Del: {item.DeliveryTime} Weeks";
                lstItems.Items.Add(itemText);
            }
            lblItemCount.Text = $"Total Items: {rfqItems.Count}";
            btnEditItem.Enabled = false;
            btnRemoveItem.Enabled = false;
        }

        private void ClearItemForm()
        {
            txtItemDescription.Clear();
            numItemQuantity.Value = 1;
            txtItemUnit.Text = "PCS";
            numItemUnitPrice.Value = 0;
            numItemDeliveryTime.Value = 0;
        }

        private bool ValidateItemForm()
        {
            if (string.IsNullOrWhiteSpace(txtItemDescription.Text))
            {
                MessageBox.Show("Please enter a description for this item.",
                    "Description Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtItemDescription.Focus();
                return false;
            }
            if (numItemQuantity.Value <= 0)
            {
                MessageBox.Show("Quantity must be at least 1.",
                    "Invalid Quantity", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                numItemQuantity.Focus();
                return false;
            }
            if (string.IsNullOrWhiteSpace(txtItemUnit.Text))
            {
                MessageBox.Show("Please enter the unit (e.g. PCS, SET, KG).",
                    "Unit Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtItemUnit.Focus();
                return false;
            }
            return true;
        }
        #endregion

        #region Validation
        private bool ValidateRFQHeader()
        {
            if (cmbCompany.SelectedValue == null || (int)cmbCompany.SelectedValue == 0)
            {
                MessageBox.Show("Please select a company before continuing.",
                    "Company Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbCompany.Focus();
                return false;
            }
            if (cmbClient.SelectedValue == null || (int)cmbClient.SelectedValue == 0)
            {
                MessageBox.Show("Please select a client before continuing.",
                    "Client Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbClient.Focus();
                return false;
            }
            if (string.IsNullOrWhiteSpace(txtRFQCode.Text))
            {
                MessageBox.Show("Please enter the RFQ Code.",
                    "RFQ Code Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtRFQCode.Focus();
                return false;
            }
            if (string.IsNullOrWhiteSpace(txtQuoteCode.Text) || txtQuoteCode.Text.Contains("[Error") || txtQuoteCode.Text.Contains("[Could not"))
            {
                MessageBox.Show("The Quote Code could not be generated. Please reselect the company and client, or enter it manually.",
                    "Quote Code Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtQuoteCode.Focus();
                return false;
            }
            return true;
        }

        private bool ValidateRFQItems()
        {
            if (rfqItems.Count == 0)
            {
                MessageBox.Show("Please add at least one item before generating.",
                    "No Items", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }
        #endregion

        #region Save RFQ
        private bool SaveRFQ()
        {
            if (!ValidateRFQHeader() || !ValidateRFQItems()) return false;

            string selectedDeliveryTerm = "DAP";
            if (rbtnDAP.Checked) selectedDeliveryTerm = "DAP";
            else if (rbtnDDP.Checked) selectedDeliveryTerm = "DDP";
            else if (rbtnDDUDAP.Checked) selectedDeliveryTerm = "DDU/DAP";
            else if (rbtnPCG.Checked) selectedDeliveryTerm = "PCG";

            var rfq = new RFQ
            {
                CompanyId = (int)cmbCompany.SelectedValue,
                ClientId = (int)cmbClient.SelectedValue,
                CreatedAt = dtpCreatedAt.Value,
                RFQCode = txtRFQCode.Text.Trim(),
                QuoteCode = txtQuoteCode.Text.Trim(),
                DeliveryPoint = txtDeliveryPoint.Text.Trim(),
                DeliveryTerm = selectedDeliveryTerm,
                Validity = txtValidity.Text.Trim(),
                Discount = numDiscount.Value,
                Currency = cmbCurrency.SelectedItem?.ToString() ?? "RM"
            };

            if (!isQuoteCodeManuallyEdited)
            {
                if (currentRFQId == 0)
                {
                    var selectedCompany = allCompanies.FirstOrDefault(c => c.Id == (int)cmbCompany.SelectedValue);
                    var selectedClient = allClients.FirstOrDefault(c => c.Id == (int)cmbClient.SelectedValue);
                    rfq.QuoteCode = rfqService.GenerateQuoteCode(
                        selectedCompany?.CompanyCode ?? "",
                        selectedClient?.ClientCode ?? "");
                }
                else
                {
                    rfq.QuoteCode = txtQuoteCode.Text.Trim();
                }
            }
            else
            {
                rfqService.UpdateSequenceFromManualEdit(rfq.CompanyId, rfq.QuoteCode);
            }

            currentRFQId = rfqService.SaveRFQ(rfq, rfqItems, currentRFQId);

            txtQuoteCode.Text = rfq.QuoteCode;
            txtQuoteCode.BackColor = Color.LightGreen;
            isRFQSaved = true;
            isQuoteCodeManuallyEdited = false;

            return true;
        }
        #endregion

        #region Helper - Get Output Folder

        /// <summary>
        /// Returns the folder to use as the default save location.
        /// If the user imported a pricesheet, that file's folder is returned first.
        /// Otherwise falls back to the network RFQ folder.
        /// </summary>
        private string GetOutputFolder()
        {
            // ✅ If the user imported a pricesheet, default to that file's folder
            if (!string.IsNullOrEmpty(importedFileFolder) && Directory.Exists(importedFileFolder))
                return importedFileFolder;

            // Otherwise use the network RFQ folder
            try
            {
                if (!Directory.Exists(RFQOutputFolder))
                    Directory.CreateDirectory(RFQOutputFolder);
                return RFQOutputFolder;
            }
            catch
            {
                string fallback = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "RFQ Generator");

                if (!Directory.Exists(fallback))
                    Directory.CreateDirectory(fallback);

                MessageBox.Show(
                    $"The network folder is not accessible right now.\n\nFiles will be saved to your Documents folder instead:\n{fallback}",
                    "Network Unavailable", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                return fallback;
            }
        }


        // ✅ Modern Windows Explorer-style folder browser dialog via COM shell (works on .NET Framework)
        private string ShowModernFolderDialog(string description, string initialPath)
        {
            try
            {
                var dialog = (IFileOpenDialog)new FileOpenDialogRCW();

                uint options;
                dialog.GetOptions(out options);
                options |= FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_NOREADONLYRETURN;
                dialog.SetOptions(options);
                dialog.SetTitle(description);

                if (!string.IsNullOrEmpty(initialPath) && Directory.Exists(initialPath))
                {
                    IShellItem item;
                    var shellItemGuid = typeof(IShellItem).GUID;
                    SHCreateItemFromParsingName(initialPath, IntPtr.Zero, ref shellItemGuid, out item);
                    if (item != null)
                        dialog.SetFolder(item);
                }

                int hr = dialog.Show(this.Handle);
                if (hr != 0) return null;

                IShellItem resultItem;
                dialog.GetResult(out resultItem);

                string path;
                resultItem.GetDisplayName(SIGDN_FILESYSPATH, out path);
                return path;
            }
            catch
            {
                using (var fallback = new FolderBrowserDialog())
                {
                    fallback.Description = description;
                    fallback.SelectedPath = initialPath;
                    fallback.ShowNewFolderButton = true;

                    if (fallback.ShowDialog() == DialogResult.OK)
                        return fallback.SelectedPath;

                    return null;
                }
            }
        }

        #region COM Interop for Modern Folder Dialog
        private const uint FOS_PICKFOLDERS = 0x00000020;
        private const uint FOS_FORCEFILESYSTEM = 0x00000040;
        private const uint FOS_NOREADONLYRETURN = 0x00000800;
        private const uint SIGDN_FILESYSPATH = 0x80058000;

        [System.Runtime.InteropServices.ComImport]
        [System.Runtime.InteropServices.Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")]
        [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
        private class FileOpenDialogRCW { }

        [System.Runtime.InteropServices.ComImport]
        [System.Runtime.InteropServices.Guid("D57C7288-D4AD-4768-BE02-9D969532D960")]
        [System.Runtime.InteropServices.InterfaceType(System.Runtime.InteropServices.ComInterfaceType.InterfaceIsIUnknown)]
        private interface IFileOpenDialog
        {
            [System.Runtime.InteropServices.PreserveSig] int Show(IntPtr hwnd);
            void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
            void SetFileTypeIndex(uint iFileType);
            void GetFileTypeIndex(out uint piFileType);
            void Advise(IntPtr pfde, out uint pdwCookie);
            void Unadvise(uint dwCookie);
            void SetOptions(uint fos);
            void GetOptions(out uint pfos);
            void SetDefaultFolder(IShellItem psi);
            void SetFolder(IShellItem psi);
            void GetFolder(out IShellItem ppsi);
            void GetCurrentSelection(out IShellItem ppsi);
            void SetFileName([System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)] string pszName);
            void GetFileName([System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)] out string pszName);
            void SetTitle([System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)] string pszTitle);
            void SetOkButtonLabel([System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)] string pszText);
            void SetFileNameLabel([System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)] string pszLabel);
            void GetResult(out IShellItem ppsi);
            void AddPlace(IShellItem psi, int fdap);
            void SetDefaultExtension([System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)] string pszDefaultExtension);
            void Close(int hr);
            void SetClientGuid(ref Guid guid);
            void ClearClientData();
            void SetFilter(IntPtr pFilter);
            void GetResults(out IntPtr ppenum);
            void GetSelectedItems(out IntPtr ppenum);
        }

        [System.Runtime.InteropServices.ComImport]
        [System.Runtime.InteropServices.Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
        [System.Runtime.InteropServices.InterfaceType(System.Runtime.InteropServices.ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellItem
        {
            void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
            void GetParent(out IShellItem ppsi);
            void GetDisplayName(uint sigdnName, [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)] out string ppszName);
            void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
            void Compare(IShellItem psi, uint hint, out int piOrder);
        }

        [System.Runtime.InteropServices.DllImport("shell32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        private static extern int SHCreateItemFromParsingName(
            string pszPath,
            IntPtr pbc,
            [System.Runtime.InteropServices.In] ref Guid riid,
            out IShellItem ppv);
        #endregion
        #endregion

        #region Generate Excel (PRICED)
        private void btnGenerateExcelPriced_Click(object sender, EventArgs e)
        {
            GenerateExcel(isPriced: true);
        }
        #endregion

        #region Generate Excel (UNPRICED)
        private void btnGenerateExcelUnpriced_Click(object sender, EventArgs e)
        {
            GenerateExcel(isPriced: false);
        }
        #endregion

        #region Generate Excel (Common Method)
        private void GenerateExcel(bool isPriced)
        {
            try
            {
                if (!ValidateRFQHeader() || !ValidateRFQItems()) return;
                if (currentTemplate == null)
                {
                    MessageBox.Show("No template found for this company. Please select a valid company.",
                        "Template Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!SaveRFQ()) return;

                var (rfq, items) = rfqService.GetRFQWithItems(currentRFQId);

                string selectedCurrency = cmbCurrency.SelectedItem?.ToString() ?? "RM";
                if (string.IsNullOrEmpty(rfq.Currency) || rfq.Currency != selectedCurrency)
                    rfq.Currency = selectedCurrency;

                string versionSuffix = excelService.GenerateRFQExcel(
                    currentTemplate.TemplatePath,
                    Path.Combine(Path.GetTempPath(), $"temp_{Guid.NewGuid()}.xlsx"),
                    rfq, items, currentTemplate.Id, isPriced);

                string outputFolder = GetOutputFolder();
                string defaultFileName = $"{rfq.QuoteCode}-{versionSuffix}.xlsx";

                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                    sfd.FileName = defaultFileName;
                    sfd.Title = $"Save {versionSuffix} RFQ Excel File";
                    sfd.InitialDirectory = outputFolder;

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        excelService.GenerateRFQExcel(
                            currentTemplate.TemplatePath,
                            sfd.FileName,
                            rfq, items, currentTemplate.Id, isPriced);

                        MessageBox.Show(
                            $"Excel file saved successfully!\n\nQuote Code: {rfq.QuoteCode}\nSaved to: {sfd.FileName}",
                            "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not generate the Excel file. Please make sure the previous generated file with the same quote code is closed before generating and try again.",
                    "Excel Generation Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region Generate BOTH Excel Versions
        private void btnGenerateBothExcel_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ValidateRFQHeader() || !ValidateRFQItems()) return;
                if (currentTemplate == null)
                {
                    MessageBox.Show("No template found for this company. Please select a valid company.",
                        "Template Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!SaveRFQ()) return;

                var (rfq, items) = rfqService.GetRFQWithItems(currentRFQId);

                string selectedCurrency = cmbCurrency.SelectedItem?.ToString() ?? "RM";
                if (string.IsNullOrEmpty(rfq.Currency) || rfq.Currency != selectedCurrency)
                    rfq.Currency = selectedCurrency;

                string selectedFolder = ShowModernFolderDialog(
                    "Select folder to save both Excel files",
                    GetOutputFolder());

                if (selectedFolder != null)
                {
                    string baseOutputPath = Path.Combine(selectedFolder, $"{rfq.QuoteCode}.xlsx");

                    var (pricedPath, unpricedPath, pricedSuffix, unpricedSuffix) =
                        excelService.GenerateBothVersionsExcel(
                            currentTemplate.TemplatePath,
                            baseOutputPath, rfq, items, currentTemplate.Id);

                    MessageBox.Show(
                        $"Both Excel files saved successfully!\n\nQuote Code: {rfq.QuoteCode}\n\n{pricedSuffix}: {pricedPath}\n{unpricedSuffix}: {unpricedPath}",
                        "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not generate the Excel files. Please make sure the template file is closed before generating and try again.",
                    "Excel Generation Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region Generate PDF (PRICED)
        private void btnGeneratePDFPriced_Click(object sender, EventArgs e)
        {
            GeneratePDF(isPriced: true);
        }
        #endregion

        #region Generate PDF (UNPRICED)
        private void btnGeneratePDFUnpriced_Click(object sender, EventArgs e)
        {
            GeneratePDF(isPriced: false);
        }
        #endregion

        #region Generate PDF (Common Method)
        private void GeneratePDF(bool isPriced)
        {
            try
            {
                if (!ValidateRFQHeader() || !ValidateRFQItems()) return;
                if (currentTemplate == null)
                {
                    MessageBox.Show("No template found for this company. Please select a valid company.",
                        "Template Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!SaveRFQ()) return;

                var (rfq, items) = rfqService.GetRFQWithItems(currentRFQId);

                string selectedCurrency = cmbCurrency.SelectedItem?.ToString() ?? "RM";
                if (string.IsNullOrEmpty(rfq.Currency) || rfq.Currency != selectedCurrency)
                    rfq.Currency = selectedCurrency;

                string versionSuffix = pdfService.GenerateRFQPDF(
                    currentTemplate.TemplatePath,
                    Path.Combine(Path.GetTempPath(), $"temp_{Guid.NewGuid()}.pdf"),
                    rfq, items, currentTemplate.Id, isPriced);

                string outputFolder = GetOutputFolder();
                string defaultFileName = $"{rfq.QuoteCode}-{versionSuffix}.pdf";

                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*";
                    sfd.FileName = defaultFileName;
                    sfd.Title = $"Save {versionSuffix} RFQ PDF File";
                    sfd.InitialDirectory = outputFolder;

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        pdfService.GenerateRFQPDF(
                            currentTemplate.TemplatePath,
                            sfd.FileName,
                            rfq, items, currentTemplate.Id, isPriced);

                        MessageBox.Show(
                            $"PDF file saved successfully!\n\nQuote Code: {rfq.QuoteCode}\nSaved to: {sfd.FileName}",
                            "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not generate the PDF file. Please make sure the template is accessible and try again.",
                    "PDF Generation Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region Generate BOTH PDF Versions
        private void btnGenerateBothPDF_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ValidateRFQHeader() || !ValidateRFQItems()) return;
                if (currentTemplate == null)
                {
                    MessageBox.Show("No template found for this company. Please select a valid company.",
                        "Template Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!SaveRFQ()) return;

                var (rfq, items) = rfqService.GetRFQWithItems(currentRFQId);

                string selectedCurrency = cmbCurrency.SelectedItem?.ToString() ?? "RM";
                if (string.IsNullOrEmpty(rfq.Currency) || rfq.Currency != selectedCurrency)
                    rfq.Currency = selectedCurrency;

                string selectedFolder = ShowModernFolderDialog(
                    "Select folder to save both PDF files",
                    GetOutputFolder());

                if (selectedFolder != null)
                {
                    string baseOutputPath = Path.Combine(selectedFolder, $"{rfq.QuoteCode}.pdf");

                    var (pricedPath, unpricedPath, pricedSuffix, unpricedSuffix) =
                        pdfService.GenerateBothVersionsPDF(
                            currentTemplate.TemplatePath,
                            baseOutputPath, rfq, items, currentTemplate.Id);

                    MessageBox.Show(
                        $"Both PDF files saved successfully!\n\nQuote Code: {rfq.QuoteCode}\n\n{pricedSuffix}: {pricedPath}\n{unpricedSuffix}: {unpricedPath}",
                        "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not generate the PDF files. Please make sure the template is accessible and try again.",
                    "PDF Generation Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region New RFQ
        private void btnNew_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show(
                "Start a new RFQ? Any unsaved changes will be lost.",
                "New RFQ", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
                ClearForm();
        }

        private void ClearForm()
        {
            cmbCompany.SelectedIndex = 0;
            cmbClient.SelectedIndex = 0;
            txtRFQCode.Clear();
            txtQuoteCode.Clear();
            txtQuoteCode.BackColor = Color.LightYellow;
            txtDeliveryPoint.Clear();
            txtValidity.Clear();
            numDiscount.Value = 0;
            cmbCurrency.SelectedIndex = 0;
            dtpCreatedAt.Value = DateTime.Now;
            rfqItems.Clear();
            UpdateItemsList();
            ClearItemForm();
            currentRFQId = 0;
            editingItemIndex = -1;
            isQuoteCodeManuallyEdited = false;
            isRFQSaved = false;
            btnAddItem.Text = "Add Item";
            importedFileFolder = null;

            rbtnDAP.Checked = true;
            rfqService.ClearQuoteCodeCache();
            cmbCompany.Focus();
        }
        #endregion

        #region Drag and Drop Implementation

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length == 1 && importService.IsValidExcelFile(files[0]))
                {
                    e.Effect = DragDropEffects.Copy;
                    return;
                }
            }
            e.Effect = DragDropEffects.None;
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            try
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                    if (files.Length != 1)
                    {
                        MessageBox.Show("Please drop only one Excel file at a time.",
                            "Too Many Files", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    string filePath = files[0];

                    if (!importService.IsValidExcelFile(filePath))
                    {
                        MessageBox.Show("Please drop a valid Excel file (.xlsx or .xls).",
                            "Invalid File", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    this.BeginInvoke(new Action(() => ImportRFQFromFile(filePath)));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not read the dropped file. Please try again.",
                    "Drop Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ImportRFQFromFile(string filePath)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;
                Application.DoEvents();

                var importedData = importService.ImportFromExcel(filePath);

                string capturedFolder = Path.GetDirectoryName(filePath);
                rfqItems.Clear();
                UpdateItemsList();
                ClearItemForm();
                editingItemIndex = -1;
                btnAddItem.Text = "Add Item";
                importedFileFolder = capturedFolder;
                if (!string.IsNullOrEmpty(importedData.DeliveryPoint))
                    txtDeliveryPoint.Text = importedData.DeliveryPoint;

                if (!string.IsNullOrEmpty(importedData.Validity))
                    txtValidity.Text = importedData.Validity;

                if (importedData.Items.Count > 0)
                {
                    foreach (var importedItem in importedData.Items)
                    {
                        txtItemDescription.Text = importedItem.ItemDesc;
                        numItemQuantity.Value = importedItem.Quantity;
                        txtItemUnit.Text = importedItem.Unit;
                        numItemUnitPrice.Value = (decimal)importedItem.UnitPrice;
                        numItemDeliveryTime.Value = importedItem.DeliveryTime;
                        btnAddItem_Click(null, null);
                    }
                    lstItems.Focus();
                }

                this.Cursor = Cursors.Default;

                MessageBox.Show(
                    $"{importedData.Items.Count} item(s) imported successfully.",
                    "Import Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                MessageBox.Show("Could not import the file. Please make sure it is a valid pricesheet and try again.",
                    "Import Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        private void lblItemDeliveryTime_Click(object sender, EventArgs e) { }
        private void numDiscount_ValueChanged(object sender, EventArgs e) { }
    }
}