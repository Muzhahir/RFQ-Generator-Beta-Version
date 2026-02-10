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
        private Template currentTemplate;
        private int currentRFQId = 0;
        private List<RFQItem> rfqItems;
        private int editingItemIndex = -1;
        private List<Company> allCompanies;
        private List<Client> allClients;
        private bool isQuoteCodeManuallyEdited = false; // Track if user manually edited the quote code

        public Form1()
        {
            InitializeComponent();
            rfqService = new RFQService();
            excelService = new ExcelGenerationService();
            pdfService = new PDFGenerationService();
            rfqItems = new List<RFQItem>();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                // CRITICAL: Change DropDownStyle to DropDown for searchable capability
                cmbCompany.DropDownStyle = ComboBoxStyle.DropDown;
                cmbClient.DropDownStyle = ComboBoxStyle.DropDown;

                // Load companies and clients into dropdowns
                LoadCompanies();
                LoadClients();

                // Set default date
                dtpCreatedAt.Value = DateTime.Now;

                // Set default currency
                cmbCurrency.SelectedIndex = 0; // Default to RM

                // Initialize item form
                ClearItemForm();
                UpdateItemsList();

                // Make quote code textbox EDITABLE but with visual hint
                txtQuoteCode.ReadOnly = false;
                txtQuoteCode.BackColor = Color.LightYellow;

                // IMPORTANT: Generate buttons are ALWAYS enabled (no Save required)
                btnGeneratePDF.Enabled = true;           // Excel button
                btnGeneratePDFDirect.Enabled = true;     // PDF button

                // CRITICAL: Manually attach event handlers AFTER loading data
                // This ensures they fire when user makes changes
                cmbCompany.SelectedIndexChanged += HandleCompanyOrClientChange;
                cmbClient.SelectedIndexChanged += HandleCompanyOrClientChange;
                txtQuoteCode.TextChanged += txtQuoteCode_TextChanged;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading form: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region Data Loading
        private void LoadCompanies()
        {
            try
            {
                allCompanies = rfqService.GetAllCompanies();

                // Create a list with a placeholder item
                var companiesWithPlaceholder = new List<Company>();
                companiesWithPlaceholder.Add(new Company { Id = 0, CompanyName = "-- Select Company --", CompanyCode = "" });
                companiesWithPlaceholder.AddRange(allCompanies);

                cmbCompany.DataSource = companiesWithPlaceholder;
                cmbCompany.DisplayMember = "CompanyName";
                cmbCompany.ValueMember = "Id";
                cmbCompany.SelectedIndex = 0;

                // Enable autocomplete
                cmbCompany.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                cmbCompany.AutoCompleteSource = AutoCompleteSource.ListItems;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading companies: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadClients()
        {
            try
            {
                allClients = rfqService.GetAllClients();

                // Create a list with a placeholder item
                var clientsWithPlaceholder = new List<Client>();
                clientsWithPlaceholder.Add(new Client { Id = 0, ClientName = "-- Select Client --", ClientCode = "" });
                clientsWithPlaceholder.AddRange(allClients);

                cmbClient.DataSource = clientsWithPlaceholder;
                cmbClient.DisplayMember = "ClientName";
                cmbClient.ValueMember = "Id";
                cmbClient.SelectedIndex = 0;

                // Enable autocomplete
                cmbClient.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                cmbClient.AutoCompleteSource = AutoCompleteSource.ListItems;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading clients: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region Quote Code Generation
        /// <summary>
        /// Track when user manually edits the quote code
        /// </summary>
        private void txtQuoteCode_TextChanged(object sender, EventArgs e)
        {
            // If the text was changed by the user (not by our auto-generation), mark it as manually edited
            if (txtQuoteCode.Focused)
            {
                isQuoteCodeManuallyEdited = true;
                txtQuoteCode.BackColor = Color.LightGreen; // Different color to show it's custom
            }
        }

        /// <summary>
        /// Single handler for both company and client changes
        /// This ensures quote code updates whenever either dropdown changes
        /// </summary>
        private void HandleCompanyOrClientChange(object sender, EventArgs e)
        {
            // Update template if company changed
            if (sender == cmbCompany)
            {
                UpdateTemplateDisplay();
            }

            // Only update quote code preview if user hasn't manually edited it
            if (!isQuoteCodeManuallyEdited)
            {
                GenerateAndDisplayQuoteCodePreview();
            }
        }

        /// <summary>
        /// Update template display based on selected company
        /// </summary>
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
                MessageBox.Show($"Error loading template: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Generate and display quote code PREVIEW (does NOT increment sequence)
        /// </summary>
        private void GenerateAndDisplayQuoteCodePreview()
        {
            try
            {
                // Clear if either selection is invalid
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

                // Get selected company and client
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
                        txtQuoteCode.Text = "[Company code not set in database]";
                        txtQuoteCode.BackColor = Color.LightCoral;
                        MessageBox.Show($"Company '{selectedCompany.CompanyName}' does not have a CompanyCode set in the database.\n\nPlease run the UpdateCompanyClientCodes.sql script.",
                            "Missing Company Code", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Generate quote code preview
                    string quoteCode = rfqService.GenerateQuoteCodePreview(companyCode, clientCode);

                    if (string.IsNullOrEmpty(quoteCode))
                    {
                        txtQuoteCode.Text = "[Error generating code]";
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
                MessageBox.Show($"Error generating quote code preview:\n{ex.Message}\n\nStack Trace:\n{ex.StackTrace}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region Company Selection & Template Loading (FROM DESIGNER)
        /// <summary>
        /// This is called from the designer - we'll keep it but it now just calls our unified handler
        /// </summary>
        private void cmbCompany_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Designer event - redirect to our unified handler
            HandleCompanyOrClientChange(sender, e);
        }
        #endregion

        #region Item Management
        private void btnAddItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ValidateItemForm())
                    return;

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
                {
                    rfqItems[i].ItemNo = i + 1;
                }

                UpdateItemsList();
                ClearItemForm();
                txtItemDescription.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding item: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                MessageBox.Show("Please select an item to edit.", "Information",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                MessageBox.Show("Please select an item to remove.", "Information",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var result = MessageBox.Show("Are you sure you want to remove this item?",
                "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                rfqItems.RemoveAt(lstItems.SelectedIndex);

                for (int i = 0; i < rfqItems.Count; i++)
                {
                    rfqItems[i].ItemNo = i + 1;
                }

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
                if (shortDesc.Length > 40)
                    shortDesc = shortDesc.Substring(0, 40) + "...";

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
                MessageBox.Show("Please enter item description.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtItemDescription.Focus();
                return false;
            }

            if (numItemQuantity.Value <= 0)
            {
                MessageBox.Show("Quantity must be greater than zero.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                numItemQuantity.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtItemUnit.Text))
            {
                MessageBox.Show("Please enter unit name.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                MessageBox.Show("Please select a company.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbCompany.Focus();
                return false;
            }

            if (cmbClient.SelectedValue == null || (int)cmbClient.SelectedValue == 0)
            {
                MessageBox.Show("Please select a client.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbClient.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtRFQCode.Text))
            {
                MessageBox.Show("Please enter RFQ Code.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtRFQCode.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtQuoteCode.Text) || txtQuoteCode.Text.Contains("[Error"))
            {
                MessageBox.Show("Quote Code was not generated. Please reselect company and client, or enter manually.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtQuoteCode.Focus();
                return false;
            }

            return true;
        }

        private bool ValidateRFQItems()
        {
            if (rfqItems.Count == 0)
            {
                MessageBox.Show("Please add at least one item.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }
        #endregion

        #region Save RFQ (Internal - Auto-save on Generate)
        /// <summary>
        /// Internal save logic - called automatically when Generate button is clicked
        /// No separate Save button needed - everything happens on Generate
        /// Returns true if save was successful
        /// </summary>
        private bool SaveRFQ()
        {
            if (!ValidateRFQHeader() || !ValidateRFQItems())
                return false;

            var rfq = new RFQ
            {
                CompanyId = (int)cmbCompany.SelectedValue,
                ClientId = (int)cmbClient.SelectedValue,
                CreatedAt = dtpCreatedAt.Value,
                RFQCode = txtRFQCode.Text.Trim(),
                QuoteCode = txtQuoteCode.Text.Trim(),
                DeliveryPoint = txtDeliveryPoint.Text.Trim(),
                DeliveryTerm = cmbDeliveryTerm.Text.Trim(),
                Validity = txtValidity.Text.Trim(),
                Discount = numDiscount.Value,
                Currency = cmbCurrency.SelectedItem?.ToString() ?? "RM"
            };

            // If user manually edited the quote code, use it directly
            // Otherwise, generate a new one with sequence increment
            if (!isQuoteCodeManuallyEdited)
            {
                var selectedCompany = allCompanies.FirstOrDefault(c => c.Id == (int)cmbCompany.SelectedValue);
                var selectedClient = allClients.FirstOrDefault(c => c.Id == (int)cmbClient.SelectedValue);
                string companyCode = selectedCompany?.CompanyCode ?? "";
                string clientCode = selectedClient?.ClientCode ?? "";

                // Generate the actual quote code with sequence increment
                rfq.QuoteCode = rfqService.GenerateQuoteCode(companyCode, clientCode);
            }

            // Save RFQ (quote code is already set)
            currentRFQId = rfqService.SaveRFQ(rfq, rfqItems);

            // Update the display with the saved quote code
            txtQuoteCode.Text = rfq.QuoteCode;
            txtQuoteCode.BackColor = Color.LightGreen;

            // Reset the manual edit flag for next RFQ
            isQuoteCodeManuallyEdited = false;

            return true;
        }
        #endregion

        #region Generate Excel (with Auto-Save)
        /// <summary>
        /// Generate Excel - AUTOMATICALLY SAVES the RFQ first, then generates the Excel file
        /// This is the ONLY button users need to click - no separate Save button required
        /// </summary>
        private void btnGenerateExcel_Click(object sender, EventArgs e)
        {
            try
            {
                // Validate first
                if (!ValidateRFQHeader() || !ValidateRFQItems())
                    return;

                if (currentTemplate == null)
                {
                    MessageBox.Show("No template available for the selected company.", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // AUTO-SAVE: Always save the RFQ (new or update)
                if (!SaveRFQ())
                {
                    return; // Save failed, don't continue to generate
                }

                // Now generate the Excel file using the saved RFQ
                var (rfq, items) = rfqService.GetRFQWithItems(currentRFQId);

                // CRITICAL FIX: Ensure currency from form is used (in case database retrieval has issues)
                string selectedCurrency = cmbCurrency.SelectedItem?.ToString() ?? "RM";
                if (string.IsNullOrEmpty(rfq.Currency) || rfq.Currency != selectedCurrency)
                {
                    // Force currency from form if database value is wrong
                    rfq.Currency = selectedCurrency;
                }

                string defaultFileName = $"RFQ_{rfq.RFQCode}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                    sfd.FileName = defaultFileName;
                    sfd.Title = "Save RFQ Excel File";

                    string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    string rfqFolder = Path.Combine(documentsPath, "RFQ Generator");
                    if (!Directory.Exists(rfqFolder))
                        Directory.CreateDirectory(rfqFolder);

                    sfd.InitialDirectory = rfqFolder;

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        excelService.GenerateRFQExcel(
                            currentTemplate.TemplatePath,
                            sfd.FileName,
                            rfq,
                            items,
                            currentTemplate.Id
                        );

                        MessageBox.Show($"Success! RFQ saved to database and Excel file generated.\n\nRFQ ID: {currentRFQId}\nQuote Code: {rfq.QuoteCode}\nCurrency: {rfq.Currency}\n\nFile saved to:\n{sfd.FileName}",
                            "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        var result = MessageBox.Show("Do you want to open the generated file?",
                            "Open File", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (result == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start(sfd.FileName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating Excel file:\n\n{ex.Message}\n\nStack Trace:\n{ex.StackTrace}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region Generate PDF (Direct Generation)
        /// <summary>
        /// Generate PDF directly from current form data without saving to database
        /// </summary>
        private void btnGeneratePDF_Click(object sender, EventArgs e)
        {
            try
            {
                // Validate first
                if (!ValidateRFQHeader() || !ValidateRFQItems())
                    return;

                if (currentTemplate == null)
                {
                    MessageBox.Show("No template available for the selected company.", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var rfq = new RFQ
                {
                    CompanyId = (int)cmbCompany.SelectedValue,
                    ClientId = (int)cmbClient.SelectedValue,
                    CreatedAt = dtpCreatedAt.Value,
                    RFQCode = txtRFQCode.Text.Trim(),
                    QuoteCode = txtQuoteCode.Text.Trim(),
                    DeliveryPoint = txtDeliveryPoint.Text.Trim(),
                    DeliveryTerm = cmbDeliveryTerm.Text.Trim(),
                    Validity = txtValidity.Text.Trim(),
                    Discount = numDiscount.Value,
                    Currency = cmbCurrency.SelectedItem?.ToString() ?? "RM"
                };

                string defaultFileName = $"RFQ_{rfq.RFQCode}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";

                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*";
                    sfd.FileName = defaultFileName;
                    sfd.Title = "Save RFQ PDF File";

                    string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    string rfqFolder = Path.Combine(documentsPath, "RFQ Generator");
                    if (!Directory.Exists(rfqFolder))
                        Directory.CreateDirectory(rfqFolder);

                    sfd.InitialDirectory = rfqFolder;

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        // Generate PDF directly from form data
                        pdfService.GenerateRFQPDF(
                            currentTemplate.TemplatePath,
                            sfd.FileName,
                            rfq,
                            rfqItems,
                            currentTemplate.Id
                        );

                        MessageBox.Show($"Success! PDF file generated.\n\nFile saved to:\n{sfd.FileName}",
                            "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        var result = MessageBox.Show("Do you want to open the generated file?",
                            "Open File", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (result == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start(sfd.FileName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error generating PDF file:\n\n{ex.Message}\n\nStack Trace:\n{ex.StackTrace}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region New RFQ
        private void btnNew_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show(
                "Are you sure you want to create a new RFQ? Any unsaved changes will be lost.",
                "Confirm New RFQ",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                ClearForm();
            }
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
            cmbCurrency.SelectedIndex = 0; // Reset to RM
            dtpCreatedAt.Value = DateTime.Now;
            rfqItems.Clear();
            UpdateItemsList();
            ClearItemForm();
            currentRFQId = 0;
            editingItemIndex = -1;
            isQuoteCodeManuallyEdited = false;
            btnAddItem.Text = "Add Item";
            btnGeneratePDF.Enabled = true;
            btnGeneratePDFDirect.Enabled = true;

            cmbCompany.Focus();
        }
        #endregion
    }
}