namespace RFQ_Generator_System
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.grpHeader = new System.Windows.Forms.GroupBox();
            this.lblDeliveryTerm = new System.Windows.Forms.Label();
            this.rbtnDAP = new System.Windows.Forms.RadioButton();
            this.rbtnDDP = new System.Windows.Forms.RadioButton();
            this.numDiscount = new System.Windows.Forms.NumericUpDown();
            this.lblDiscount = new System.Windows.Forms.Label();
            this.txtValidity = new System.Windows.Forms.TextBox();
            this.lblValidity = new System.Windows.Forms.Label();
            this.dtpCreatedAt = new System.Windows.Forms.DateTimePicker();
            this.lblCreatedAt = new System.Windows.Forms.Label();
            this.txtDeliveryPoint = new System.Windows.Forms.TextBox();
            this.lblDeliveryPoint = new System.Windows.Forms.Label();
            this.txtQuoteCode = new System.Windows.Forms.TextBox();
            this.lblQuoteCode = new System.Windows.Forms.Label();
            this.txtRFQCode = new System.Windows.Forms.TextBox();
            this.lblRFQCode = new System.Windows.Forms.Label();
            this.cmbClient = new System.Windows.Forms.ComboBox();
            this.lblClient = new System.Windows.Forms.Label();
            this.cmbCompany = new System.Windows.Forms.ComboBox();
            this.lblCompany = new System.Windows.Forms.Label();
            this.cmbCurrency = new System.Windows.Forms.ComboBox();
            this.lblCurrency = new System.Windows.Forms.Label();
            this.grpCurrentItem = new System.Windows.Forms.GroupBox();
            this.numItemDeliveryTime = new System.Windows.Forms.NumericUpDown();
            this.lblItemDeliveryTime = new System.Windows.Forms.Label();
            this.numItemUnitPrice = new System.Windows.Forms.NumericUpDown();
            this.lblItemUnitPrice = new System.Windows.Forms.Label();
            this.txtItemUnit = new System.Windows.Forms.TextBox();
            this.lblItemUnit = new System.Windows.Forms.Label();
            this.numItemQuantity = new System.Windows.Forms.NumericUpDown();
            this.lblItemQuantity = new System.Windows.Forms.Label();
            this.txtItemDescription = new System.Windows.Forms.TextBox();
            this.lblItemDescription = new System.Windows.Forms.Label();
            this.btnAddItem = new System.Windows.Forms.Button();
            this.btnClearItem = new System.Windows.Forms.Button();
            this.grpItemsList = new System.Windows.Forms.GroupBox();
            this.lstItems = new System.Windows.Forms.ListBox();
            this.btnRemoveItem = new System.Windows.Forms.Button();
            this.btnEditItem = new System.Windows.Forms.Button();
            this.lblItemCount = new System.Windows.Forms.Label();
            this.grpActions = new System.Windows.Forms.GroupBox();
            this.btnGeneratePDFUnpriced = new System.Windows.Forms.Button();
            this.btnGeneratePDFPriced = new System.Windows.Forms.Button();
            this.btnGenerateExcelUnpriced = new System.Windows.Forms.Button();
            this.btnGenerateExcelPriced = new System.Windows.Forms.Button();
            this.lblTemplatePath = new System.Windows.Forms.Label();
            this.lblTemplateLabel = new System.Windows.Forms.Label();
            this.btnNew = new System.Windows.Forms.Button();
            this.grpHeader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numDiscount)).BeginInit();
            this.grpCurrentItem.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numItemDeliveryTime)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numItemUnitPrice)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numItemQuantity)).BeginInit();
            this.grpItemsList.SuspendLayout();
            this.grpActions.SuspendLayout();
            this.SuspendLayout();
            // 
            // grpHeader
            // 
            this.grpHeader.Controls.Add(this.lblDeliveryTerm);
            this.grpHeader.Controls.Add(this.rbtnDAP);
            this.grpHeader.Controls.Add(this.rbtnDDP);
            this.grpHeader.Controls.Add(this.numDiscount);
            this.grpHeader.Controls.Add(this.lblDiscount);
            this.grpHeader.Controls.Add(this.txtValidity);
            this.grpHeader.Controls.Add(this.lblValidity);
            this.grpHeader.Controls.Add(this.dtpCreatedAt);
            this.grpHeader.Controls.Add(this.lblCreatedAt);
            this.grpHeader.Controls.Add(this.txtDeliveryPoint);
            this.grpHeader.Controls.Add(this.lblDeliveryPoint);
            this.grpHeader.Controls.Add(this.txtQuoteCode);
            this.grpHeader.Controls.Add(this.lblQuoteCode);
            this.grpHeader.Controls.Add(this.txtRFQCode);
            this.grpHeader.Controls.Add(this.lblRFQCode);
            this.grpHeader.Controls.Add(this.cmbClient);
            this.grpHeader.Controls.Add(this.lblClient);
            this.grpHeader.Controls.Add(this.cmbCompany);
            this.grpHeader.Controls.Add(this.lblCompany);
            this.grpHeader.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.grpHeader.Location = new System.Drawing.Point(16, 15);
            this.grpHeader.Margin = new System.Windows.Forms.Padding(4);
            this.grpHeader.Name = "grpHeader";
            this.grpHeader.Padding = new System.Windows.Forms.Padding(4);
            this.grpHeader.Size = new System.Drawing.Size(1413, 258);
            this.grpHeader.TabIndex = 0;
            this.grpHeader.TabStop = false;
            this.grpHeader.Text = "RFQ Header Information";
            // 
            // lblDeliveryTerm
            // 
            this.lblDeliveryTerm.AutoSize = true;
            this.lblDeliveryTerm.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblDeliveryTerm.Location = new System.Drawing.Point(40, 178);
            this.lblDeliveryTerm.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblDeliveryTerm.Name = "lblDeliveryTerm";
            this.lblDeliveryTerm.Size = new System.Drawing.Size(103, 20);
            this.lblDeliveryTerm.TabIndex = 12;
            this.lblDeliveryTerm.Text = "Delivery Term:";
            // 
            // rbtnDAP
            // 
            this.rbtnDAP.AutoSize = true;
            this.rbtnDAP.Checked = true;
            this.rbtnDAP.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.rbtnDAP.Location = new System.Drawing.Point(40, 205);
            this.rbtnDAP.Margin = new System.Windows.Forms.Padding(4);
            this.rbtnDAP.Name = "rbtnDAP";
            this.rbtnDAP.Size = new System.Drawing.Size(59, 24);
            this.rbtnDAP.TabIndex = 20;
            this.rbtnDAP.TabStop = true;
            this.rbtnDAP.Text = "DAP";
            this.rbtnDAP.UseVisualStyleBackColor = true;
            // 
            // rbtnDDP
            // 
            this.rbtnDDP.AutoSize = true;
            this.rbtnDDP.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.rbtnDDP.Location = new System.Drawing.Point(130, 205);
            this.rbtnDDP.Margin = new System.Windows.Forms.Padding(4);
            this.rbtnDDP.Name = "rbtnDDP";
            this.rbtnDDP.Size = new System.Drawing.Size(60, 24);
            this.rbtnDDP.TabIndex = 21;
            this.rbtnDDP.Text = "DDP";
            this.rbtnDDP.UseVisualStyleBackColor = true;
            // 
            // numDiscount
            // 
            this.numDiscount.DecimalPlaces = 2;
            this.numDiscount.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.numDiscount.Location = new System.Drawing.Point(867, 203);
            this.numDiscount.Margin = new System.Windows.Forms.Padding(4);
            this.numDiscount.Maximum = new decimal(new int[] {
            1000000,
            0,
            0,
            0});
            this.numDiscount.Name = "numDiscount";
            this.numDiscount.Size = new System.Drawing.Size(507, 27);
            this.numDiscount.TabIndex = 17;
            // 
            // lblDiscount
            // 
            this.lblDiscount.AutoSize = true;
            this.lblDiscount.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblDiscount.Location = new System.Drawing.Point(867, 178);
            this.lblDiscount.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblDiscount.Name = "lblDiscount";
            this.lblDiscount.Size = new System.Drawing.Size(96, 20);
            this.lblDiscount.TabIndex = 16;
            this.lblDiscount.Text = "Discount (%):";
            // 
            // txtValidity
            // 
            this.txtValidity.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtValidity.Location = new System.Drawing.Point(1097, 50);
            this.txtValidity.Margin = new System.Windows.Forms.Padding(4);
            this.txtValidity.Name = "txtValidity";
            this.txtValidity.Size = new System.Drawing.Size(277, 27);
            this.txtValidity.TabIndex = 15;
            // 
            // lblValidity
            // 
            this.lblValidity.AutoSize = true;
            this.lblValidity.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblValidity.Location = new System.Drawing.Point(1093, 26);
            this.lblValidity.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblValidity.Name = "lblValidity";
            this.lblValidity.Size = new System.Drawing.Size(107, 20);
            this.lblValidity.TabIndex = 14;
            this.lblValidity.Text = "Validity (Days):";
            // 
            // dtpCreatedAt
            // 
            this.dtpCreatedAt.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.dtpCreatedAt.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpCreatedAt.Location = new System.Drawing.Point(867, 117);
            this.dtpCreatedAt.Margin = new System.Windows.Forms.Padding(4);
            this.dtpCreatedAt.Name = "dtpCreatedAt";
            this.dtpCreatedAt.Size = new System.Drawing.Size(505, 27);
            this.dtpCreatedAt.TabIndex = 11;
            // 
            // lblCreatedAt
            // 
            this.lblCreatedAt.AutoSize = true;
            this.lblCreatedAt.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblCreatedAt.Location = new System.Drawing.Point(867, 92);
            this.lblCreatedAt.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblCreatedAt.Name = "lblCreatedAt";
            this.lblCreatedAt.Size = new System.Drawing.Size(100, 20);
            this.lblCreatedAt.TabIndex = 10;
            this.lblCreatedAt.Text = "Created Date:";
            // 
            // txtDeliveryPoint
            // 
            this.txtDeliveryPoint.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtDeliveryPoint.Location = new System.Drawing.Point(453, 202);
            this.txtDeliveryPoint.Margin = new System.Windows.Forms.Padding(4);
            this.txtDeliveryPoint.Name = "txtDeliveryPoint";
            this.txtDeliveryPoint.Size = new System.Drawing.Size(372, 27);
            this.txtDeliveryPoint.TabIndex = 9;
            // 
            // lblDeliveryPoint
            // 
            this.lblDeliveryPoint.AutoSize = true;
            this.lblDeliveryPoint.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblDeliveryPoint.Location = new System.Drawing.Point(453, 178);
            this.lblDeliveryPoint.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblDeliveryPoint.Name = "lblDeliveryPoint";
            this.lblDeliveryPoint.Size = new System.Drawing.Size(103, 20);
            this.lblDeliveryPoint.TabIndex = 8;
            this.lblDeliveryPoint.Text = "Delivery Point:";
            // 
            // txtQuoteCode
            // 
            this.txtQuoteCode.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtQuoteCode.Location = new System.Drawing.Point(453, 117);
            this.txtQuoteCode.Margin = new System.Windows.Forms.Padding(4);
            this.txtQuoteCode.Name = "txtQuoteCode";
            this.txtQuoteCode.Size = new System.Drawing.Size(372, 27);
            this.txtQuoteCode.TabIndex = 7;
            // 
            // lblQuoteCode
            // 
            this.lblQuoteCode.AutoSize = true;
            this.lblQuoteCode.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblQuoteCode.Location = new System.Drawing.Point(453, 92);
            this.lblQuoteCode.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblQuoteCode.Name = "lblQuoteCode";
            this.lblQuoteCode.Size = new System.Drawing.Size(92, 20);
            this.lblQuoteCode.TabIndex = 6;
            this.lblQuoteCode.Text = "Quote Code:";
            // 
            // txtRFQCode
            // 
            this.txtRFQCode.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtRFQCode.Location = new System.Drawing.Point(40, 117);
            this.txtRFQCode.Margin = new System.Windows.Forms.Padding(4);
            this.txtRFQCode.Name = "txtRFQCode";
            this.txtRFQCode.Size = new System.Drawing.Size(372, 27);
            this.txtRFQCode.TabIndex = 5;
            // 
            // lblRFQCode
            // 
            this.lblRFQCode.AutoSize = true;
            this.lblRFQCode.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblRFQCode.Location = new System.Drawing.Point(40, 92);
            this.lblRFQCode.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblRFQCode.Name = "lblRFQCode";
            this.lblRFQCode.Size = new System.Drawing.Size(78, 20);
            this.lblRFQCode.TabIndex = 4;
            this.lblRFQCode.Text = "RFQ Code:";
            // 
            // cmbClient
            // 
            this.cmbClient.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.cmbClient.FormattingEnabled = true;
            this.cmbClient.Location = new System.Drawing.Point(453, 49);
            this.cmbClient.Margin = new System.Windows.Forms.Padding(4);
            this.cmbClient.Name = "cmbClient";
            this.cmbClient.Size = new System.Drawing.Size(604, 28);
            this.cmbClient.TabIndex = 3;
            // 
            // lblClient
            // 
            this.lblClient.AutoSize = true;
            this.lblClient.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblClient.Location = new System.Drawing.Point(453, 25);
            this.lblClient.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblClient.Name = "lblClient";
            this.lblClient.Size = new System.Drawing.Size(50, 20);
            this.lblClient.TabIndex = 2;
            this.lblClient.Text = "Client:";
            // 
            // cmbCompany
            // 
            this.cmbCompany.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.cmbCompany.FormattingEnabled = true;
            this.cmbCompany.Location = new System.Drawing.Point(40, 49);
            this.cmbCompany.Margin = new System.Windows.Forms.Padding(4);
            this.cmbCompany.Name = "cmbCompany";
            this.cmbCompany.Size = new System.Drawing.Size(372, 28);
            this.cmbCompany.TabIndex = 1;
            this.cmbCompany.SelectedIndexChanged += new System.EventHandler(this.cmbCompany_SelectedIndexChanged);
            // 
            // lblCompany
            // 
            this.lblCompany.AutoSize = true;
            this.lblCompany.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblCompany.Location = new System.Drawing.Point(40, 25);
            this.lblCompany.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblCompany.Name = "lblCompany";
            this.lblCompany.Size = new System.Drawing.Size(75, 20);
            this.lblCompany.TabIndex = 0;
            this.lblCompany.Text = "Company:";
            // 
            // cmbCurrency
            // 
            this.cmbCurrency.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbCurrency.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.cmbCurrency.FormattingEnabled = true;
            this.cmbCurrency.Items.AddRange(new object[] {
            "RM",
            "$",
            "SGD",
            "€",
            "£",
            "¥"});
            this.cmbCurrency.Location = new System.Drawing.Point(469, 191);
            this.cmbCurrency.Margin = new System.Windows.Forms.Padding(4);
            this.cmbCurrency.Name = "cmbCurrency";
            this.cmbCurrency.Size = new System.Drawing.Size(86, 28);
            this.cmbCurrency.TabIndex = 19;
            // 
            // lblCurrency
            // 
            this.lblCurrency.AutoSize = true;
            this.lblCurrency.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblCurrency.Location = new System.Drawing.Point(465, 166);
            this.lblCurrency.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblCurrency.Name = "lblCurrency";
            this.lblCurrency.Size = new System.Drawing.Size(69, 20);
            this.lblCurrency.TabIndex = 18;
            this.lblCurrency.Text = "Currency:";
            // 
            // grpCurrentItem
            // 
            this.grpCurrentItem.Controls.Add(this.lblCurrency);
            this.grpCurrentItem.Controls.Add(this.cmbCurrency);
            this.grpCurrentItem.Controls.Add(this.numItemDeliveryTime);
            this.grpCurrentItem.Controls.Add(this.lblItemDeliveryTime);
            this.grpCurrentItem.Controls.Add(this.numItemUnitPrice);
            this.grpCurrentItem.Controls.Add(this.lblItemUnitPrice);
            this.grpCurrentItem.Controls.Add(this.txtItemUnit);
            this.grpCurrentItem.Controls.Add(this.lblItemUnit);
            this.grpCurrentItem.Controls.Add(this.numItemQuantity);
            this.grpCurrentItem.Controls.Add(this.lblItemQuantity);
            this.grpCurrentItem.Controls.Add(this.txtItemDescription);
            this.grpCurrentItem.Controls.Add(this.lblItemDescription);
            this.grpCurrentItem.Controls.Add(this.btnAddItem);
            this.grpCurrentItem.Controls.Add(this.btnClearItem);
            this.grpCurrentItem.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.grpCurrentItem.Location = new System.Drawing.Point(16, 281);
            this.grpCurrentItem.Margin = new System.Windows.Forms.Padding(4);
            this.grpCurrentItem.Name = "grpCurrentItem";
            this.grpCurrentItem.Padding = new System.Windows.Forms.Padding(4);
            this.grpCurrentItem.Size = new System.Drawing.Size(748, 308);
            this.grpCurrentItem.TabIndex = 1;
            this.grpCurrentItem.TabStop = false;
            this.grpCurrentItem.Text = "Item Details";
            // 
            // numItemDeliveryTime
            // 
            this.numItemDeliveryTime.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.numItemDeliveryTime.Location = new System.Drawing.Point(572, 191);
            this.numItemDeliveryTime.Margin = new System.Windows.Forms.Padding(4);
            this.numItemDeliveryTime.Maximum = new decimal(new int[] {
            365,
            0,
            0,
            0});
            this.numItemDeliveryTime.Name = "numItemDeliveryTime";
            this.numItemDeliveryTime.Size = new System.Drawing.Size(146, 27);
            this.numItemDeliveryTime.TabIndex = 11;
            // 
            // lblItemDeliveryTime
            // 
            this.lblItemDeliveryTime.AutoSize = true;
            this.lblItemDeliveryTime.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblItemDeliveryTime.Location = new System.Drawing.Point(565, 167);
            this.lblItemDeliveryTime.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblItemDeliveryTime.Name = "lblItemDeliveryTime";
            this.lblItemDeliveryTime.Size = new System.Drawing.Size(159, 20);
            this.lblItemDeliveryTime.TabIndex = 10;
            this.lblItemDeliveryTime.Text = "Delivery Time (Weeks):";
            // 
            // numItemUnitPrice
            // 
            this.numItemUnitPrice.DecimalPlaces = 2;
            this.numItemUnitPrice.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.numItemUnitPrice.Location = new System.Drawing.Point(279, 191);
            this.numItemUnitPrice.Margin = new System.Windows.Forms.Padding(4);
            this.numItemUnitPrice.Maximum = new decimal(new int[] {
            1000000,
            0,
            0,
            0});
            this.numItemUnitPrice.Name = "numItemUnitPrice";
            this.numItemUnitPrice.Size = new System.Drawing.Size(173, 27);
            this.numItemUnitPrice.TabIndex = 9;
            // 
            // lblItemUnitPrice
            // 
            this.lblItemUnitPrice.AutoSize = true;
            this.lblItemUnitPrice.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblItemUnitPrice.Location = new System.Drawing.Point(275, 167);
            this.lblItemUnitPrice.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblItemUnitPrice.Name = "lblItemUnitPrice";
            this.lblItemUnitPrice.Size = new System.Drawing.Size(75, 20);
            this.lblItemUnitPrice.TabIndex = 8;
            this.lblItemUnitPrice.Text = "Unit Price:";
            // 
            // txtItemUnit
            // 
            this.txtItemUnit.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtItemUnit.Location = new System.Drawing.Point(192, 191);
            this.txtItemUnit.Margin = new System.Windows.Forms.Padding(4);
            this.txtItemUnit.Name = "txtItemUnit";
            this.txtItemUnit.Size = new System.Drawing.Size(69, 27);
            this.txtItemUnit.TabIndex = 7;
            this.txtItemUnit.Text = "PCS";
            // 
            // lblItemUnit
            // 
            this.lblItemUnit.AutoSize = true;
            this.lblItemUnit.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblItemUnit.Location = new System.Drawing.Point(188, 167);
            this.lblItemUnit.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblItemUnit.Name = "lblItemUnit";
            this.lblItemUnit.Size = new System.Drawing.Size(39, 20);
            this.lblItemUnit.TabIndex = 6;
            this.lblItemUnit.Text = "Unit:";
            // 
            // numItemQuantity
            // 
            this.numItemQuantity.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.numItemQuantity.Location = new System.Drawing.Point(40, 191);
            this.numItemQuantity.Margin = new System.Windows.Forms.Padding(4);
            this.numItemQuantity.Maximum = new decimal(new int[] {
            100000,
            0,
            0,
            0});
            this.numItemQuantity.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numItemQuantity.Name = "numItemQuantity";
            this.numItemQuantity.Size = new System.Drawing.Size(133, 27);
            this.numItemQuantity.TabIndex = 5;
            this.numItemQuantity.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // lblItemQuantity
            // 
            this.lblItemQuantity.AutoSize = true;
            this.lblItemQuantity.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblItemQuantity.Location = new System.Drawing.Point(40, 166);
            this.lblItemQuantity.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblItemQuantity.Name = "lblItemQuantity";
            this.lblItemQuantity.Size = new System.Drawing.Size(68, 20);
            this.lblItemQuantity.TabIndex = 4;
            this.lblItemQuantity.Text = "Quantity:";
            // 
            // txtItemDescription
            // 
            this.txtItemDescription.AcceptsReturn = true;
            this.txtItemDescription.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtItemDescription.Location = new System.Drawing.Point(40, 68);
            this.txtItemDescription.Margin = new System.Windows.Forms.Padding(4);
            this.txtItemDescription.Multiline = true;
            this.txtItemDescription.Name = "txtItemDescription";
            this.txtItemDescription.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtItemDescription.Size = new System.Drawing.Size(678, 79);
            this.txtItemDescription.TabIndex = 3;
            // 
            // lblItemDescription
            // 
            this.lblItemDescription.AutoSize = true;
            this.lblItemDescription.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblItemDescription.Location = new System.Drawing.Point(40, 43);
            this.lblItemDescription.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblItemDescription.Name = "lblItemDescription";
            this.lblItemDescription.Size = new System.Drawing.Size(88, 20);
            this.lblItemDescription.TabIndex = 2;
            this.lblItemDescription.Text = "Description:";
            // 
            // btnAddItem
            // 
            this.btnAddItem.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnAddItem.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnAddItem.Location = new System.Drawing.Point(40, 246);
            this.btnAddItem.Margin = new System.Windows.Forms.Padding(4);
            this.btnAddItem.Name = "btnAddItem";
            this.btnAddItem.Size = new System.Drawing.Size(160, 37);
            this.btnAddItem.TabIndex = 0;
            this.btnAddItem.Text = "Add Item";
            this.btnAddItem.UseVisualStyleBackColor = false;
            this.btnAddItem.Click += new System.EventHandler(this.btnAddItem_Click);
            // 
            // btnClearItem
            // 
            this.btnClearItem.BackColor = System.Drawing.Color.DarkSalmon;
            this.btnClearItem.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnClearItem.Location = new System.Drawing.Point(213, 246);
            this.btnClearItem.Margin = new System.Windows.Forms.Padding(4);
            this.btnClearItem.Name = "btnClearItem";
            this.btnClearItem.Size = new System.Drawing.Size(160, 37);
            this.btnClearItem.TabIndex = 1;
            this.btnClearItem.Text = "Clear";
            this.btnClearItem.UseVisualStyleBackColor = false;
            this.btnClearItem.Click += new System.EventHandler(this.btnClearItem_Click);
            // 
            // grpItemsList
            // 
            this.grpItemsList.Controls.Add(this.lstItems);
            this.grpItemsList.Controls.Add(this.btnRemoveItem);
            this.grpItemsList.Controls.Add(this.btnEditItem);
            this.grpItemsList.Controls.Add(this.lblItemCount);
            this.grpItemsList.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.grpItemsList.Location = new System.Drawing.Point(793, 281);
            this.grpItemsList.Margin = new System.Windows.Forms.Padding(4);
            this.grpItemsList.Name = "grpItemsList";
            this.grpItemsList.Padding = new System.Windows.Forms.Padding(4);
            this.grpItemsList.Size = new System.Drawing.Size(637, 308);
            this.grpItemsList.TabIndex = 2;
            this.grpItemsList.TabStop = false;
            this.grpItemsList.Text = "Items List";
            // 
            // lstItems
            // 
            this.lstItems.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstItems.FormattingEnabled = true;
            this.lstItems.ItemHeight = 17;
            this.lstItems.Location = new System.Drawing.Point(27, 43);
            this.lstItems.Margin = new System.Windows.Forms.Padding(4);
            this.lstItems.Name = "lstItems";
            this.lstItems.Size = new System.Drawing.Size(570, 191);
            this.lstItems.TabIndex = 0;
            this.lstItems.SelectedIndexChanged += new System.EventHandler(this.lstItems_SelectedIndexChanged);
            // 
            // btnRemoveItem
            // 
            this.btnRemoveItem.BackColor = System.Drawing.Color.DarkSalmon;
            this.btnRemoveItem.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnRemoveItem.Location = new System.Drawing.Point(187, 258);
            this.btnRemoveItem.Margin = new System.Windows.Forms.Padding(4);
            this.btnRemoveItem.Name = "btnRemoveItem";
            this.btnRemoveItem.Size = new System.Drawing.Size(133, 37);
            this.btnRemoveItem.TabIndex = 2;
            this.btnRemoveItem.Text = "Remove";
            this.btnRemoveItem.UseVisualStyleBackColor = false;
            this.btnRemoveItem.Click += new System.EventHandler(this.btnRemoveItem_Click);
            // 
            // btnEditItem
            // 
            this.btnEditItem.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnEditItem.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnEditItem.Location = new System.Drawing.Point(27, 258);
            this.btnEditItem.Margin = new System.Windows.Forms.Padding(4);
            this.btnEditItem.Name = "btnEditItem";
            this.btnEditItem.Size = new System.Drawing.Size(133, 37);
            this.btnEditItem.TabIndex = 1;
            this.btnEditItem.Text = "Edit";
            this.btnEditItem.UseVisualStyleBackColor = false;
            this.btnEditItem.Click += new System.EventHandler(this.btnEditItem_Click);
            // 
            // lblItemCount
            // 
            this.lblItemCount.AutoSize = true;
            this.lblItemCount.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblItemCount.ForeColor = System.Drawing.Color.Gray;
            this.lblItemCount.Location = new System.Drawing.Point(333, 265);
            this.lblItemCount.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblItemCount.Name = "lblItemCount";
            this.lblItemCount.Size = new System.Drawing.Size(91, 19);
            this.lblItemCount.TabIndex = 3;
            this.lblItemCount.Text = "Total Items: 0";
            // 
            // grpActions
            // 
            this.grpActions.Controls.Add(this.btnGeneratePDFUnpriced);
            this.grpActions.Controls.Add(this.btnGeneratePDFPriced);
            this.grpActions.Controls.Add(this.btnGenerateExcelUnpriced);
            this.grpActions.Controls.Add(this.btnGenerateExcelPriced);
            this.grpActions.Controls.Add(this.lblTemplatePath);
            this.grpActions.Controls.Add(this.lblTemplateLabel);
            this.grpActions.Controls.Add(this.btnNew);
            this.grpActions.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.grpActions.Location = new System.Drawing.Point(16, 596);
            this.grpActions.Margin = new System.Windows.Forms.Padding(4);
            this.grpActions.Name = "grpActions";
            this.grpActions.Padding = new System.Windows.Forms.Padding(4);
            this.grpActions.Size = new System.Drawing.Size(1413, 150);
            this.grpActions.TabIndex = 3;
            this.grpActions.TabStop = false;
            this.grpActions.Text = "Actions";
            // 
            // btnGeneratePDFUnpriced
            // 
            this.btnGeneratePDFUnpriced.BackColor = System.Drawing.Color.Turquoise;
            this.btnGeneratePDFUnpriced.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnGeneratePDFUnpriced.Location = new System.Drawing.Point(821, 31);
            this.btnGeneratePDFUnpriced.Margin = new System.Windows.Forms.Padding(4);
            this.btnGeneratePDFUnpriced.Name = "btnGeneratePDFUnpriced";
            this.btnGeneratePDFUnpriced.Size = new System.Drawing.Size(200, 37);
            this.btnGeneratePDFUnpriced.TabIndex = 8;
            this.btnGeneratePDFUnpriced.Text = "PDF Unpriced/Technical";
            this.btnGeneratePDFUnpriced.UseVisualStyleBackColor = false;
            this.btnGeneratePDFUnpriced.Click += new System.EventHandler(this.btnGeneratePDFUnpriced_Click);
            // 
            // btnGeneratePDFPriced
            // 
            this.btnGeneratePDFPriced.BackColor = System.Drawing.Color.Turquoise;
            this.btnGeneratePDFPriced.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnGeneratePDFPriced.Location = new System.Drawing.Point(613, 31);
            this.btnGeneratePDFPriced.Margin = new System.Windows.Forms.Padding(4);
            this.btnGeneratePDFPriced.Name = "btnGeneratePDFPriced";
            this.btnGeneratePDFPriced.Size = new System.Drawing.Size(200, 37);
            this.btnGeneratePDFPriced.TabIndex = 7;
            this.btnGeneratePDFPriced.Text = "PDF Priced/Commercial";
            this.btnGeneratePDFPriced.UseVisualStyleBackColor = false;
            this.btnGeneratePDFPriced.Click += new System.EventHandler(this.btnGeneratePDFPriced_Click);
            // 
            // btnGenerateExcelUnpriced
            // 
            this.btnGenerateExcelUnpriced.BackColor = System.Drawing.Color.LightGreen;
            this.btnGenerateExcelUnpriced.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnGenerateExcelUnpriced.Location = new System.Drawing.Point(405, 31);
            this.btnGenerateExcelUnpriced.Margin = new System.Windows.Forms.Padding(4);
            this.btnGenerateExcelUnpriced.Name = "btnGenerateExcelUnpriced";
            this.btnGenerateExcelUnpriced.Size = new System.Drawing.Size(200, 37);
            this.btnGenerateExcelUnpriced.TabIndex = 6;
            this.btnGenerateExcelUnpriced.Text = "Excel Unpriced/Technical";
            this.btnGenerateExcelUnpriced.UseVisualStyleBackColor = false;
            this.btnGenerateExcelUnpriced.Click += new System.EventHandler(this.btnGenerateExcelUnpriced_Click);
            // 
            // btnGenerateExcelPriced
            // 
            this.btnGenerateExcelPriced.BackColor = System.Drawing.Color.LightGreen;
            this.btnGenerateExcelPriced.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnGenerateExcelPriced.Location = new System.Drawing.Point(181, 31);
            this.btnGenerateExcelPriced.Margin = new System.Windows.Forms.Padding(4);
            this.btnGenerateExcelPriced.Name = "btnGenerateExcelPriced";
            this.btnGenerateExcelPriced.Size = new System.Drawing.Size(216, 37);
            this.btnGenerateExcelPriced.TabIndex = 5;
            this.btnGenerateExcelPriced.Text = "Excel Priced/Commercial";
            this.btnGenerateExcelPriced.UseVisualStyleBackColor = false;
            this.btnGenerateExcelPriced.Click += new System.EventHandler(this.btnGenerateExcelPriced_Click);
            // 
            // lblTemplatePath
            // 
            this.lblTemplatePath.AutoSize = true;
            this.lblTemplatePath.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Italic);
            this.lblTemplatePath.ForeColor = System.Drawing.Color.Gray;
            this.lblTemplatePath.Location = new System.Drawing.Point(133, 113);
            this.lblTemplatePath.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblTemplatePath.Name = "lblTemplatePath";
            this.lblTemplatePath.Size = new System.Drawing.Size(142, 19);
            this.lblTemplatePath.TabIndex = 4;
            this.lblTemplatePath.Text = "No template selected";
            // 
            // lblTemplateLabel
            // 
            this.lblTemplateLabel.AutoSize = true;
            this.lblTemplateLabel.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblTemplateLabel.Location = new System.Drawing.Point(40, 113);
            this.lblTemplateLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblTemplateLabel.Name = "lblTemplateLabel";
            this.lblTemplateLabel.Size = new System.Drawing.Size(67, 19);
            this.lblTemplateLabel.TabIndex = 3;
            this.lblTemplateLabel.Text = "Template:";
            // 
            // btnNew
            // 
            this.btnNew.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnNew.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnNew.Location = new System.Drawing.Point(40, 31);
            this.btnNew.Margin = new System.Windows.Forms.Padding(4);
            this.btnNew.Name = "btnNew";
            this.btnNew.Size = new System.Drawing.Size(133, 37);
            this.btnNew.TabIndex = 0;
            this.btnNew.Text = "New RFQ";
            this.btnNew.UseVisualStyleBackColor = false;
            this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1445, 760);
            this.Controls.Add(this.grpActions);
            this.Controls.Add(this.grpItemsList);
            this.Controls.Add(this.grpCurrentItem);
            this.Controls.Add(this.grpHeader);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "RFQ Generator System";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.grpHeader.ResumeLayout(false);
            this.grpHeader.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numDiscount)).EndInit();
            this.grpCurrentItem.ResumeLayout(false);
            this.grpCurrentItem.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numItemDeliveryTime)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numItemUnitPrice)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numItemQuantity)).EndInit();
            this.grpItemsList.ResumeLayout(false);
            this.grpItemsList.PerformLayout();
            this.grpActions.ResumeLayout(false);
            this.grpActions.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox grpHeader;
        private System.Windows.Forms.ComboBox cmbCompany;
        private System.Windows.Forms.Label lblCompany;
        private System.Windows.Forms.ComboBox cmbClient;
        private System.Windows.Forms.Label lblClient;
        private System.Windows.Forms.TextBox txtRFQCode;
        private System.Windows.Forms.Label lblRFQCode;
        private System.Windows.Forms.TextBox txtQuoteCode;
        private System.Windows.Forms.Label lblQuoteCode;
        private System.Windows.Forms.TextBox txtDeliveryPoint;
        private System.Windows.Forms.Label lblDeliveryPoint;
        private System.Windows.Forms.DateTimePicker dtpCreatedAt;
        private System.Windows.Forms.Label lblCreatedAt;
        private System.Windows.Forms.Label lblDeliveryTerm;
        private System.Windows.Forms.RadioButton rbtnDAP;
        private System.Windows.Forms.RadioButton rbtnDDP;
        private System.Windows.Forms.TextBox txtValidity;
        private System.Windows.Forms.Label lblValidity;
        private System.Windows.Forms.NumericUpDown numDiscount;
        private System.Windows.Forms.Label lblDiscount;
        private System.Windows.Forms.ComboBox cmbCurrency;
        private System.Windows.Forms.Label lblCurrency;
        private System.Windows.Forms.GroupBox grpCurrentItem;
        private System.Windows.Forms.TextBox txtItemDescription;
        private System.Windows.Forms.Label lblItemDescription;
        private System.Windows.Forms.NumericUpDown numItemQuantity;
        private System.Windows.Forms.Label lblItemQuantity;
        private System.Windows.Forms.TextBox txtItemUnit;
        private System.Windows.Forms.Label lblItemUnit;
        private System.Windows.Forms.NumericUpDown numItemUnitPrice;
        private System.Windows.Forms.Label lblItemUnitPrice;
        private System.Windows.Forms.NumericUpDown numItemDeliveryTime;
        private System.Windows.Forms.Label lblItemDeliveryTime;
        private System.Windows.Forms.Button btnAddItem;
        private System.Windows.Forms.Button btnClearItem;
        private System.Windows.Forms.GroupBox grpItemsList;
        private System.Windows.Forms.ListBox lstItems;
        private System.Windows.Forms.Button btnRemoveItem;
        private System.Windows.Forms.Button btnEditItem;
        private System.Windows.Forms.Label lblItemCount;
        private System.Windows.Forms.GroupBox grpActions;
        private System.Windows.Forms.Button btnNew;
        private System.Windows.Forms.Button btnGenerateExcelPriced;
        private System.Windows.Forms.Button btnGenerateExcelUnpriced;
        private System.Windows.Forms.Button btnGeneratePDFPriced;
        private System.Windows.Forms.Button btnGeneratePDFUnpriced;
        private System.Windows.Forms.Label lblTemplatePath;
        private System.Windows.Forms.Label lblTemplateLabel;
    }
}