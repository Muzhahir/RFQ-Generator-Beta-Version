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
        private Template currentTemplate;
        private int currentRFQId = 0;
        private List<RFQItem> rfqItems;
        private int editingItemIndex = -1; // Track which item is being edited

        public Form1()
        {
            InitializeComponent();
            rfqService = new RFQService();
            excelService = new ExcelGenerationService();
            rfqItems = new List<RFQItem>();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                // Load companies and clients into dropdowns
                LoadCompanies();
                LoadClients();

                // Set default date
                dtpCreatedAt.Value = DateTime.Now;

                // Initialize item form
                ClearItemForm();
                UpdateItemsList();
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
                var companies = rfqService.GetAllCompanies();

                cmbCompany.DataSource = companies;
                cmbCompany.DisplayMember = "CompanyName";
                cmbCompany.ValueMember = "Id";

                if (companies.Count > 0)
                {
                    cmbCompany.SelectedIndex = 0;
                }
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
                var clients = rfqService.GetAllClients();

                cmbClient.DataSource = clients;
                cmbClient.DisplayMember = "ClientName";
                cmbClient.ValueMember = "Id";

                if (clients.Count > 0)
                {
                    cmbClient.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading clients: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Company Selection & Template Loading

        private void cmbCompany_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbCompany.SelectedValue == null || !(cmbCompany.SelectedValue is int))
                return;

            try
            {
                int companyId = (int)cmbCompany.SelectedValue;

                // Load template for selected company
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

        #endregion

        #region Item Management

        private void btnAddItem_Click(object sender, EventArgs e)
        {
            try
            {
                // Validate item fields
                if (!ValidateItemForm())
                    return;

                var item = new RFQItem
                {
                    ItemDesc = txtItemDescription.Text.Trim(),
                    Quantity = (int)numItemQuantity.Value,
                    UnitName = txtItemUnit.Text.Trim(),
                    UnitPrice = numItemUnitPrice.Value,
                    DeliveryTime = (int)numItemDeliveryTime.Value
                    // REMOVED: DeliveryTerm - it's only in the header
                };

                if (editingItemIndex >= 0)
                {
                    // Update existing item
                    rfqItems[editingItemIndex] = item;
                    editingItemIndex = -1;
                    btnAddItem.Text = "Add Item";
                }
                else
                {
                    // Add new item
                    rfqItems.Add(item);
                }

                // Renumber items
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

            // Load item data into form
            txtItemDescription.Text = item.ItemDesc;
            numItemQuantity.Value = item.Quantity;
            txtItemUnit.Text = item.UnitName;
            numItemUnitPrice.Value = item.UnitPrice;
            numItemDeliveryTime.Value = item.DeliveryTime;
            // REMOVED: txtItemDeliveryTerm.Text = item.DeliveryTerm;

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

                // Renumber items
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
            // Enable/disable edit and remove buttons based on selection
            bool hasSelection = lstItems.SelectedIndex >= 0;
            btnEditItem.Enabled = hasSelection;
            btnRemoveItem.Enabled = hasSelection;
        }

        private void UpdateItemsList()
        {
            lstItems.Items.Clear();

            foreach (var item in rfqItems)
            {
                // Replace line breaks with space for display and truncate if too long
                string shortDesc = item.ItemDesc.Replace("\r\n", " ").Replace("\n", " ");
                if (shortDesc.Length > 40)
                    shortDesc = shortDesc.Substring(0, 40) + "...";

                string itemText = $"{item.ItemNo}. {shortDesc} | Qty: {item.Quantity} {item.UnitName} | " +
                                  $"Price: {item.UnitPrice:N2} | Del: {item.DeliveryTime} days";
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
            numItemDeliveryTime.Value = 30;
            // REMOVED: txtItemDeliveryTerm.Clear();
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
            if (cmbCompany.SelectedValue == null)
            {
                MessageBox.Show("Please select a company.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbCompany.Focus();
                return false;
            }

            if (cmbClient.SelectedValue == null)
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

            if (string.IsNullOrWhiteSpace(txtQuoteCode.Text))
            {
                MessageBox.Show("Please enter Quote Code.", "Validation Error",
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

        #region Save RFQ

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                // Validate input
                if (!ValidateRFQHeader() || !ValidateRFQItems())
                    return;

                // Create RFQ header
                var rfq = new RFQ
                {
                    CompanyId = (int)cmbCompany.SelectedValue,
                    ClientId = (int)cmbClient.SelectedValue,
                    CreatedAt = dtpCreatedAt.Value,
                    RFQCode = txtRFQCode.Text.Trim(),
                    QuoteCode = txtQuoteCode.Text.Trim(),
                    DeliveryPoint = txtDeliveryPoint.Text.Trim(),
                    DeliveryTerm = txtDeliveryTerm.Text.Trim(),
                    Validity = txtValidity.Text.Trim(),
                    Discount = numDiscount.Value
                };

                // Save using service (transaction)
                currentRFQId = rfqService.SaveRFQ(rfq, rfqItems);

                MessageBox.Show($"RFQ saved successfully!\nRFQ ID: {currentRFQId}", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Enable Generate PDF button
                btnGeneratePDF.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving RFQ: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Generate PDF/Excel

        private void btnGeneratePDF_Click(object sender, EventArgs e)
        {
            try
            {
                if (currentTemplate == null)
                {
                    MessageBox.Show("No template available for the selected company.", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (currentRFQId == 0)
                {
                    MessageBox.Show("Please save the RFQ first before generating Excel.", "Information",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Check if template file exists
                if (!File.Exists(currentTemplate.TemplatePath))
                {
                    MessageBox.Show($"Template file not found:\n{currentTemplate.TemplatePath}\n\nPlease update the template path in the database.",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Get RFQ data
                var (rfq, items) = rfqService.GetRFQWithItems(currentRFQId);

                // Create default filename
                string defaultFileName = $"RFQ_{rfq.RFQCode}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

                // Show SaveFileDialog
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                    sfd.FileName = defaultFileName;
                    sfd.Title = "Save RFQ Excel File";

                    // Set initial directory to user's Documents folder
                    string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    string rfqFolder = Path.Combine(documentsPath, "RFQ Generator");

                    // Create folder if it doesn't exist
                    if (!Directory.Exists(rfqFolder))
                        Directory.CreateDirectory(rfqFolder);

                    sfd.InitialDirectory = rfqFolder;

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        // Generate Excel file
                        excelService.GenerateRFQExcel(
                            currentTemplate.TemplatePath,
                            sfd.FileName,
                            rfq,
                            items,
                            currentTemplate.Id
                        );

                        MessageBox.Show($"RFQ Excel file generated successfully!\n\nSaved to:\n{sfd.FileName}",
                            "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Ask if user wants to open the file
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
            // Reset header fields
            if (cmbCompany.Items.Count > 0)
                cmbCompany.SelectedIndex = 0;
            if (cmbClient.Items.Count > 0)
                cmbClient.SelectedIndex = 0;

            txtRFQCode.Clear();
            txtQuoteCode.Clear();
            txtDeliveryPoint.Clear();
            txtDeliveryTerm.Clear();
            txtValidity.Clear();
            numDiscount.Value = 0;
            dtpCreatedAt.Value = DateTime.Now;

            // Clear items
            rfqItems.Clear();
            UpdateItemsList();
            ClearItemForm();

            // Reset state
            currentRFQId = 0;
            editingItemIndex = -1;
            btnGeneratePDF.Enabled = false;
            btnAddItem.Text = "Add Item";

            // Focus on first input
            txtRFQCode.Focus();
        }

        #endregion
    }
}