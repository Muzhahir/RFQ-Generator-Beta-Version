namespace RFQ_Generator_System
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.grpHeader = new System.Windows.Forms.GroupBox();
            this.numDiscount = new System.Windows.Forms.NumericUpDown();
            this.lblDiscount = new System.Windows.Forms.Label();
            this.txtValidity = new System.Windows.Forms.TextBox();
            this.lblValidity = new System.Windows.Forms.Label();
            this.txtDeliveryTerm = new System.Windows.Forms.TextBox();
            this.lblDeliveryTerm = new System.Windows.Forms.Label();
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
            this.lblTemplatePath = new System.Windows.Forms.Label();
            this.lblTemplateLabel = new System.Windows.Forms.Label();
            this.btnGeneratePDF = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
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
            this.grpHeader.Controls.Add(this.numDiscount);
            this.grpHeader.Controls.Add(this.lblDiscount);
            this.grpHeader.Controls.Add(this.txtValidity);
            this.grpHeader.Controls.Add(this.lblValidity);
            this.grpHeader.Controls.Add(this.txtDeliveryTerm);
            this.grpHeader.Controls.Add(this.lblDeliveryTerm);
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
            this.grpHeader.Location = new System.Drawing.Point(12, 12);
            this.grpHeader.Name = "grpHeader";
            this.grpHeader.Size = new System.Drawing.Size(1060, 210);
            this.grpHeader.TabIndex = 0;
            this.grpHeader.TabStop = false;
            this.grpHeader.Text = "RFQ Header Information";
            // 
            // numDiscount
            // 
            this.numDiscount.DecimalPlaces = 2;
            this.numDiscount.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.numDiscount.Location = new System.Drawing.Point(650, 165);
            this.numDiscount.Maximum = new decimal(new int[] { 1000000, 0, 0, 0 });
            this.numDiscount.Name = "numDiscount";
            this.numDiscount.Size = new System.Drawing.Size(380, 23);
            this.numDiscount.TabIndex = 17;
            // 
            // lblDiscount
            // 
            this.lblDiscount.AutoSize = true;
            this.lblDiscount.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblDiscount.Location = new System.Drawing.Point(650, 145);
            this.lblDiscount.Name = "lblDiscount";
            this.lblDiscount.Size = new System.Drawing.Size(57, 15);
            this.lblDiscount.TabIndex = 16;
            this.lblDiscount.Text = "Discount:";
            // 
            // txtValidity
            // 
            this.txtValidity.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtValidity.Location = new System.Drawing.Point(340, 165);
            this.txtValidity.Name = "txtValidity";
            this.txtValidity.Size = new System.Drawing.Size(280, 23);
            this.txtValidity.TabIndex = 15;
            // 
            // lblValidity
            // 
            this.lblValidity.AutoSize = true;
            this.lblValidity.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblValidity.Location = new System.Drawing.Point(340, 145);
            this.lblValidity.Name = "lblValidity";
            this.lblValidity.Size = new System.Drawing.Size(49, 15);
            this.lblValidity.TabIndex = 14;
            this.lblValidity.Text = "Validity:";
            // 
            // txtDeliveryTerm
            // 
            this.txtDeliveryTerm.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtDeliveryTerm.Location = new System.Drawing.Point(30, 165);
            this.txtDeliveryTerm.Name = "txtDeliveryTerm";
            this.txtDeliveryTerm.Size = new System.Drawing.Size(280, 23);
            this.txtDeliveryTerm.TabIndex = 13;
            // 
            // lblDeliveryTerm
            // 
            this.lblDeliveryTerm.AutoSize = true;
            this.lblDeliveryTerm.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblDeliveryTerm.Location = new System.Drawing.Point(30, 145);
            this.lblDeliveryTerm.Name = "lblDeliveryTerm";
            this.lblDeliveryTerm.Size = new System.Drawing.Size(84, 15);
            this.lblDeliveryTerm.TabIndex = 12;
            this.lblDeliveryTerm.Text = "Delivery Term:";
            // 
            // dtpCreatedAt
            // 
            this.dtpCreatedAt.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.dtpCreatedAt.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpCreatedAt.Location = new System.Drawing.Point(650, 95);
            this.dtpCreatedAt.Name = "dtpCreatedAt";
            this.dtpCreatedAt.Size = new System.Drawing.Size(380, 23);
            this.dtpCreatedAt.TabIndex = 11;
            // 
            // lblCreatedAt
            // 
            this.lblCreatedAt.AutoSize = true;
            this.lblCreatedAt.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblCreatedAt.Location = new System.Drawing.Point(650, 75);
            this.lblCreatedAt.Name = "lblCreatedAt";
            this.lblCreatedAt.Size = new System.Drawing.Size(76, 15);
            this.lblCreatedAt.TabIndex = 10;
            this.lblCreatedAt.Text = "Created Date:";
            // 
            // txtDeliveryPoint
            // 
            this.txtDeliveryPoint.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtDeliveryPoint.Location = new System.Drawing.Point(650, 40);
            this.txtDeliveryPoint.Name = "txtDeliveryPoint";
            this.txtDeliveryPoint.Size = new System.Drawing.Size(380, 23);
            this.txtDeliveryPoint.TabIndex = 9;
            // 
            // lblDeliveryPoint
            // 
            this.lblDeliveryPoint.AutoSize = true;
            this.lblDeliveryPoint.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblDeliveryPoint.Location = new System.Drawing.Point(650, 20);
            this.lblDeliveryPoint.Name = "lblDeliveryPoint";
            this.lblDeliveryPoint.Size = new System.Drawing.Size(85, 15);
            this.lblDeliveryPoint.TabIndex = 8;
            this.lblDeliveryPoint.Text = "Delivery Point:";
            // 
            // txtQuoteCode
            // 
            this.txtQuoteCode.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtQuoteCode.Location = new System.Drawing.Point(340, 95);
            this.txtQuoteCode.Name = "txtQuoteCode";
            this.txtQuoteCode.Size = new System.Drawing.Size(280, 23);
            this.txtQuoteCode.TabIndex = 7;
            // 
            // lblQuoteCode
            // 
            this.lblQuoteCode.AutoSize = true;
            this.lblQuoteCode.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblQuoteCode.Location = new System.Drawing.Point(340, 75);
            this.lblQuoteCode.Name = "lblQuoteCode";
            this.lblQuoteCode.Size = new System.Drawing.Size(75, 15);
            this.lblQuoteCode.TabIndex = 6;
            this.lblQuoteCode.Text = "Quote Code:";
            // 
            // txtRFQCode
            // 
            this.txtRFQCode.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtRFQCode.Location = new System.Drawing.Point(30, 95);
            this.txtRFQCode.Name = "txtRFQCode";
            this.txtRFQCode.Size = new System.Drawing.Size(280, 23);
            this.txtRFQCode.TabIndex = 5;
            // 
            // lblRFQCode
            // 
            this.lblRFQCode.AutoSize = true;
            this.lblRFQCode.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblRFQCode.Location = new System.Drawing.Point(30, 75);
            this.lblRFQCode.Name = "lblRFQCode";
            this.lblRFQCode.Size = new System.Drawing.Size(66, 15);
            this.lblRFQCode.TabIndex = 4;
            this.lblRFQCode.Text = "RFQ Code:";
            // 
            // cmbClient
            // 
            this.cmbClient.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbClient.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.cmbClient.FormattingEnabled = true;
            this.cmbClient.Location = new System.Drawing.Point(340, 40);
            this.cmbClient.Name = "cmbClient";
            this.cmbClient.Size = new System.Drawing.Size(280, 23);
            this.cmbClient.TabIndex = 3;
            // 
            // lblClient
            // 
            this.lblClient.AutoSize = true;
            this.lblClient.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblClient.Location = new System.Drawing.Point(340, 20);
            this.lblClient.Name = "lblClient";
            this.lblClient.Size = new System.Drawing.Size(41, 15);
            this.lblClient.TabIndex = 2;
            this.lblClient.Text = "Client:";
            // 
            // cmbCompany
            // 
            this.cmbCompany.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbCompany.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.cmbCompany.FormattingEnabled = true;
            this.cmbCompany.Location = new System.Drawing.Point(30, 40);
            this.cmbCompany.Name = "cmbCompany";
            this.cmbCompany.Size = new System.Drawing.Size(280, 23);
            this.cmbCompany.TabIndex = 1;
            this.cmbCompany.SelectedIndexChanged += new System.EventHandler(this.cmbCompany_SelectedIndexChanged);
            // 
            // lblCompany
            // 
            this.lblCompany.AutoSize = true;
            this.lblCompany.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblCompany.Location = new System.Drawing.Point(30, 20);
            this.lblCompany.Name = "lblCompany";
            this.lblCompany.Size = new System.Drawing.Size(62, 15);
            this.lblCompany.TabIndex = 0;
            this.lblCompany.Text = "Company:";
            // 
            // grpCurrentItem
            // 
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
            this.grpCurrentItem.Location = new System.Drawing.Point(12, 228);
            this.grpCurrentItem.Name = "grpCurrentItem";
            this.grpCurrentItem.Size = new System.Drawing.Size(680, 250);
            this.grpCurrentItem.TabIndex = 1;
            this.grpCurrentItem.TabStop = false;
            this.grpCurrentItem.Text = "Item Details";
            // 
            // numItemDeliveryTime
            // 
            this.numItemDeliveryTime.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.numItemDeliveryTime.Location = new System.Drawing.Point(490, 155);
            this.numItemDeliveryTime.Maximum = new decimal(new int[] { 365, 0, 0, 0 });
            this.numItemDeliveryTime.Name = "numItemDeliveryTime";
            this.numItemDeliveryTime.Size = new System.Drawing.Size(160, 23);
            this.numItemDeliveryTime.TabIndex = 11;
            this.numItemDeliveryTime.Value = new decimal(new int[] { 30, 0, 0, 0 });
            // 
            // lblItemDeliveryTime
            // 
            this.lblItemDeliveryTime.AutoSize = true;
            this.lblItemDeliveryTime.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblItemDeliveryTime.Location = new System.Drawing.Point(490, 135);
            this.lblItemDeliveryTime.Name = "lblItemDeliveryTime";
            this.lblItemDeliveryTime.Size = new System.Drawing.Size(124, 15);
            this.lblItemDeliveryTime.TabIndex = 10;
            this.lblItemDeliveryTime.Text = "Delivery Time (Weeks):";
            // 
            // numItemUnitPrice
            // 
            this.numItemUnitPrice.DecimalPlaces = 2;
            this.numItemUnitPrice.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.numItemUnitPrice.Location = new System.Drawing.Point(340, 155);
            this.numItemUnitPrice.Maximum = new decimal(new int[] { 1000000, 0, 0, 0 });
            this.numItemUnitPrice.Name = "numItemUnitPrice";
            this.numItemUnitPrice.Size = new System.Drawing.Size(130, 23);
            this.numItemUnitPrice.TabIndex = 9;
            // 
            // lblItemUnitPrice
            // 
            this.lblItemUnitPrice.AutoSize = true;
            this.lblItemUnitPrice.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblItemUnitPrice.Location = new System.Drawing.Point(340, 135);
            this.lblItemUnitPrice.Name = "lblItemUnitPrice";
            this.lblItemUnitPrice.Size = new System.Drawing.Size(62, 15);
            this.lblItemUnitPrice.TabIndex = 8;
            this.lblItemUnitPrice.Text = "Unit Price:";
            // 
            // txtItemUnit
            // 
            this.txtItemUnit.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtItemUnit.Location = new System.Drawing.Point(200, 155);
            this.txtItemUnit.Name = "txtItemUnit";
            this.txtItemUnit.Size = new System.Drawing.Size(120, 23);
            this.txtItemUnit.TabIndex = 7;
            this.txtItemUnit.Text = "PCS";
            // 
            // lblItemUnit
            // 
            this.lblItemUnit.AutoSize = true;
            this.lblItemUnit.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblItemUnit.Location = new System.Drawing.Point(200, 135);
            this.lblItemUnit.Name = "lblItemUnit";
            this.lblItemUnit.Size = new System.Drawing.Size(32, 15);
            this.lblItemUnit.TabIndex = 6;
            this.lblItemUnit.Text = "Unit:";
            // 
            // numItemQuantity
            // 
            this.numItemQuantity.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.numItemQuantity.Location = new System.Drawing.Point(30, 155);
            this.numItemQuantity.Maximum = new decimal(new int[] { 100000, 0, 0, 0 });
            this.numItemQuantity.Minimum = new decimal(new int[] { 1, 0, 0, 0 });
            this.numItemQuantity.Name = "numItemQuantity";
            this.numItemQuantity.Size = new System.Drawing.Size(150, 23);
            this.numItemQuantity.TabIndex = 5;
            this.numItemQuantity.Value = new decimal(new int[] { 1, 0, 0, 0 });
            // 
            // lblItemQuantity
            // 
            this.lblItemQuantity.AutoSize = true;
            this.lblItemQuantity.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblItemQuantity.Location = new System.Drawing.Point(30, 135);
            this.lblItemQuantity.Name = "lblItemQuantity";
            this.lblItemQuantity.Size = new System.Drawing.Size(56, 15);
            this.lblItemQuantity.TabIndex = 4;
            this.lblItemQuantity.Text = "Quantity:";
            // 
            // txtItemDescription
            // 
            this.txtItemDescription.AcceptsReturn = true;
            this.txtItemDescription.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtItemDescription.Location = new System.Drawing.Point(30, 55);
            this.txtItemDescription.Multiline = true;
            this.txtItemDescription.Name = "txtItemDescription";
            this.txtItemDescription.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtItemDescription.Size = new System.Drawing.Size(620, 65);
            this.txtItemDescription.TabIndex = 3;
            // 
            // lblItemDescription
            // 
            this.lblItemDescription.AutoSize = true;
            this.lblItemDescription.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblItemDescription.Location = new System.Drawing.Point(30, 35);
            this.lblItemDescription.Name = "lblItemDescription";
            this.lblItemDescription.Size = new System.Drawing.Size(70, 15);
            this.lblItemDescription.TabIndex = 2;
            this.lblItemDescription.Text = "Description:";
            // 
            // btnAddItem
            // 
            this.btnAddItem.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnAddItem.Location = new System.Drawing.Point(30, 200);
            this.btnAddItem.Name = "btnAddItem";
            this.btnAddItem.Size = new System.Drawing.Size(120, 30);
            this.btnAddItem.TabIndex = 0;
            this.btnAddItem.Text = "Add Item";
            this.btnAddItem.UseVisualStyleBackColor = true;
            this.btnAddItem.Click += new System.EventHandler(this.btnAddItem_Click);
            // 
            // btnClearItem
            // 
            this.btnClearItem.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnClearItem.Location = new System.Drawing.Point(160, 200);
            this.btnClearItem.Name = "btnClearItem";
            this.btnClearItem.Size = new System.Drawing.Size(120, 30);
            this.btnClearItem.TabIndex = 1;
            this.btnClearItem.Text = "Clear";
            this.btnClearItem.UseVisualStyleBackColor = true;
            this.btnClearItem.Click += new System.EventHandler(this.btnClearItem_Click);
            // 
            // grpItemsList
            // 
            this.grpItemsList.Controls.Add(this.lstItems);
            this.grpItemsList.Controls.Add(this.btnRemoveItem);
            this.grpItemsList.Controls.Add(this.btnEditItem);
            this.grpItemsList.Controls.Add(this.lblItemCount);
            this.grpItemsList.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.grpItemsList.Location = new System.Drawing.Point(698, 228);
            this.grpItemsList.Name = "grpItemsList";
            this.grpItemsList.Size = new System.Drawing.Size(374, 250);
            this.grpItemsList.TabIndex = 2;
            this.grpItemsList.TabStop = false;
            this.grpItemsList.Text = "Items List";
            // 
            // lstItems
            // 
            this.lstItems.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstItems.FormattingEnabled = true;
            this.lstItems.ItemHeight = 13;
            this.lstItems.Location = new System.Drawing.Point(20, 35);
            this.lstItems.Name = "lstItems";
            this.lstItems.Size = new System.Drawing.Size(334, 160);
            this.lstItems.TabIndex = 0;
            this.lstItems.SelectedIndexChanged += new System.EventHandler(this.lstItems_SelectedIndexChanged);
            // 
            // btnRemoveItem
            // 
            this.btnRemoveItem.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnRemoveItem.Location = new System.Drawing.Point(140, 210);
            this.btnRemoveItem.Name = "btnRemoveItem";
            this.btnRemoveItem.Size = new System.Drawing.Size(100, 30);
            this.btnRemoveItem.TabIndex = 2;
            this.btnRemoveItem.Text = "Remove";
            this.btnRemoveItem.UseVisualStyleBackColor = true;
            this.btnRemoveItem.Click += new System.EventHandler(this.btnRemoveItem_Click);
            // 
            // btnEditItem
            // 
            this.btnEditItem.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnEditItem.Location = new System.Drawing.Point(20, 210);
            this.btnEditItem.Name = "btnEditItem";
            this.btnEditItem.Size = new System.Drawing.Size(100, 30);
            this.btnEditItem.TabIndex = 1;
            this.btnEditItem.Text = "Edit";
            this.btnEditItem.UseVisualStyleBackColor = true;
            this.btnEditItem.Click += new System.EventHandler(this.btnEditItem_Click);
            // 
            // lblItemCount
            // 
            this.lblItemCount.AutoSize = true;
            this.lblItemCount.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblItemCount.ForeColor = System.Drawing.Color.Gray;
            this.lblItemCount.Location = new System.Drawing.Point(250, 215);
            this.lblItemCount.Name = "lblItemCount";
            this.lblItemCount.Size = new System.Drawing.Size(82, 13);
            this.lblItemCount.TabIndex = 3;
            this.lblItemCount.Text = "Total Items: 0";
            // 
            // grpActions
            // 
            this.grpActions.Controls.Add(this.lblTemplatePath);
            this.grpActions.Controls.Add(this.lblTemplateLabel);
            this.grpActions.Controls.Add(this.btnGeneratePDF);
            this.grpActions.Controls.Add(this.btnSave);
            this.grpActions.Controls.Add(this.btnNew);
            this.grpActions.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.grpActions.Location = new System.Drawing.Point(12, 484);
            this.grpActions.Name = "grpActions";
            this.grpActions.Size = new System.Drawing.Size(1060, 90);
            this.grpActions.TabIndex = 3;
            this.grpActions.TabStop = false;
            this.grpActions.Text = "Actions";
            // 
            // lblTemplatePath
            // 
            this.lblTemplatePath.AutoSize = true;
            this.lblTemplatePath.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Italic);
            this.lblTemplatePath.ForeColor = System.Drawing.Color.Gray;
            this.lblTemplatePath.Location = new System.Drawing.Point(100, 60);
            this.lblTemplatePath.Name = "lblTemplatePath";
            this.lblTemplatePath.Size = new System.Drawing.Size(120, 13);
            this.lblTemplatePath.TabIndex = 4;
            this.lblTemplatePath.Text = "No template selected";
            // 
            // lblTemplateLabel
            // 
            this.lblTemplateLabel.AutoSize = true;
            this.lblTemplateLabel.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblTemplateLabel.Location = new System.Drawing.Point(30, 60);
            this.lblTemplateLabel.Name = "lblTemplateLabel";
            this.lblTemplateLabel.Size = new System.Drawing.Size(58, 13);
            this.lblTemplateLabel.TabIndex = 3;
            this.lblTemplateLabel.Text = "Template:";
            // 
            // btnGeneratePDF
            // 
            this.btnGeneratePDF.Enabled = false;
            this.btnGeneratePDF.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnGeneratePDF.Location = new System.Drawing.Point(260, 25);
            this.btnGeneratePDF.Name = "btnGeneratePDF";
            this.btnGeneratePDF.Size = new System.Drawing.Size(120, 30);
            this.btnGeneratePDF.TabIndex = 2;
            this.btnGeneratePDF.Text = "Generate Excel";
            this.btnGeneratePDF.UseVisualStyleBackColor = true;
            this.btnGeneratePDF.Click += new System.EventHandler(this.btnGeneratePDF_Click);
            // 
            // btnSave
            // 
            this.btnSave.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnSave.Location = new System.Drawing.Point(145, 25);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(100, 30);
            this.btnSave.TabIndex = 1;
            this.btnSave.Text = "Save RFQ";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnNew
            // 
            this.btnNew.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnNew.Location = new System.Drawing.Point(30, 25);
            this.btnNew.Name = "btnNew";
            this.btnNew.Size = new System.Drawing.Size(100, 30);
            this.btnNew.TabIndex = 0;
            this.btnNew.Text = "New RFQ";
            this.btnNew.UseVisualStyleBackColor = true;
            this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1084, 586);
            this.Controls.Add(this.grpActions);
            this.Controls.Add(this.grpItemsList);
            this.Controls.Add(this.grpCurrentItem);
            this.Controls.Add(this.grpHeader);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "RFQ Generator System - CEKAP GAGASAN";
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
        private System.Windows.Forms.TextBox txtDeliveryTerm;
        private System.Windows.Forms.Label lblDeliveryTerm;
        private System.Windows.Forms.TextBox txtValidity;
        private System.Windows.Forms.Label lblValidity;
        private System.Windows.Forms.NumericUpDown numDiscount;
        private System.Windows.Forms.Label lblDiscount;
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
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnNew;
        private System.Windows.Forms.Button btnGeneratePDF;
        private System.Windows.Forms.Label lblTemplatePath;
        private System.Windows.Forms.Label lblTemplateLabel;
    }
}