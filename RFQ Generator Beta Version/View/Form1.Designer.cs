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
            this.tlpHeader = new System.Windows.Forms.TableLayoutPanel();
            this.pnlCompany = new System.Windows.Forms.Panel();
            this.cmbCompany = new System.Windows.Forms.ComboBox();
            this.lblCompany = new System.Windows.Forms.Label();
            this.pnlClient = new System.Windows.Forms.Panel();
            this.cmbClient = new System.Windows.Forms.ComboBox();
            this.lblClient = new System.Windows.Forms.Label();
            this.pnlRFQCode = new System.Windows.Forms.Panel();
            this.txtRFQCode = new System.Windows.Forms.TextBox();
            this.lblRFQCode = new System.Windows.Forms.Label();
            this.pnlQuoteCode = new System.Windows.Forms.Panel();
            this.txtQuoteCode = new System.Windows.Forms.TextBox();
            this.lblQuoteCode = new System.Windows.Forms.Label();
            this.pnlValidity = new System.Windows.Forms.Panel();
            this.txtValidity = new System.Windows.Forms.TextBox();
            this.lblValidity = new System.Windows.Forms.Label();
            this.pnlDeliveryTerm = new System.Windows.Forms.Panel();
            this.rbtnPCG = new System.Windows.Forms.RadioButton();
            this.rbtnDDUDAP = new System.Windows.Forms.RadioButton();
            this.rbtnDDP = new System.Windows.Forms.RadioButton();
            this.rbtnDAP = new System.Windows.Forms.RadioButton();
            this.lblDeliveryTerm = new System.Windows.Forms.Label();
            this.pnlDeliveryPoint = new System.Windows.Forms.Panel();
            this.txtDeliveryPoint = new System.Windows.Forms.TextBox();
            this.lblDeliveryPoint = new System.Windows.Forms.Label();
            this.pnlCreatedAt = new System.Windows.Forms.Panel();
            this.dtpCreatedAt = new System.Windows.Forms.DateTimePicker();
            this.lblCreatedAt = new System.Windows.Forms.Label();
            this.pnlDiscount = new System.Windows.Forms.Panel();
            this.numDiscount = new System.Windows.Forms.NumericUpDown();
            this.lblDiscount = new System.Windows.Forms.Label();
            this.grpCurrentItem = new System.Windows.Forms.GroupBox();
            this.txtItemDescription = new System.Windows.Forms.TextBox();
            this.lblItemDescription = new System.Windows.Forms.Label();
            this.numItemQuantity = new System.Windows.Forms.NumericUpDown();
            this.lblItemQuantity = new System.Windows.Forms.Label();
            this.txtItemUnit = new System.Windows.Forms.TextBox();
            this.lblItemUnit = new System.Windows.Forms.Label();
            this.numItemUnitPrice = new System.Windows.Forms.NumericUpDown();
            this.lblItemUnitPrice = new System.Windows.Forms.Label();
            this.cmbCurrency = new System.Windows.Forms.ComboBox();
            this.lblCurrency = new System.Windows.Forms.Label();
            this.numItemDeliveryTime = new System.Windows.Forms.NumericUpDown();
            this.lblItemDeliveryTime = new System.Windows.Forms.Label();
            this.btnAddItem = new System.Windows.Forms.Button();
            this.btnClearItem = new System.Windows.Forms.Button();
            this.grpItemsList = new System.Windows.Forms.GroupBox();
            this.lstItems = new System.Windows.Forms.ListBox();
            this.btnEditItem = new System.Windows.Forms.Button();
            this.btnRemoveItem = new System.Windows.Forms.Button();
            this.lblItemCount = new System.Windows.Forms.Label();
            this.grpActions = new System.Windows.Forms.GroupBox();
            this.btnNew = new System.Windows.Forms.Button();
            this.btnGenerateBothExcel = new System.Windows.Forms.Button();
            this.btnGenerateBothPDF = new System.Windows.Forms.Button();
            this.lblTemplateLabel = new System.Windows.Forms.Label();
            this.lblTemplatePath = new System.Windows.Forms.Label();
            this.grpHeader.SuspendLayout();
            this.tlpHeader.SuspendLayout();
            this.pnlCompany.SuspendLayout();
            this.pnlClient.SuspendLayout();
            this.pnlRFQCode.SuspendLayout();
            this.pnlQuoteCode.SuspendLayout();
            this.pnlValidity.SuspendLayout();
            this.pnlDeliveryTerm.SuspendLayout();
            this.pnlDeliveryPoint.SuspendLayout();
            this.pnlCreatedAt.SuspendLayout();
            this.pnlDiscount.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numDiscount)).BeginInit();
            this.grpCurrentItem.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numItemQuantity)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numItemUnitPrice)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numItemDeliveryTime)).BeginInit();
            this.grpItemsList.SuspendLayout();
            this.grpActions.SuspendLayout();
            this.SuspendLayout();
            // 
            // grpHeader
            // 
            this.grpHeader.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grpHeader.Controls.Add(this.tlpHeader);
            this.grpHeader.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.grpHeader.Location = new System.Drawing.Point(16, 15);
            this.grpHeader.Name = "grpHeader";
            this.grpHeader.Padding = new System.Windows.Forms.Padding(6, 4, 6, 4);
            this.grpHeader.Size = new System.Drawing.Size(1413, 155);
            this.grpHeader.TabIndex = 0;
            this.grpHeader.TabStop = false;
            this.grpHeader.Text = "RFQ Header Information";
            // 
            // tlpHeader
            // 
            this.tlpHeader.ColumnCount = 5;
            this.tlpHeader.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 21.9843F));
            this.tlpHeader.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.48537F));
            this.tlpHeader.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 21.55603F));
            this.tlpHeader.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 19.20057F));
            this.tlpHeader.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.98787F));
            this.tlpHeader.Controls.Add(this.pnlCompany, 0, 0);
            this.tlpHeader.Controls.Add(this.pnlClient, 1, 0);
            this.tlpHeader.Controls.Add(this.pnlRFQCode, 2, 0);
            this.tlpHeader.Controls.Add(this.pnlQuoteCode, 3, 0);
            this.tlpHeader.Controls.Add(this.pnlValidity, 4, 0);
            this.tlpHeader.Controls.Add(this.pnlDeliveryTerm, 0, 1);
            this.tlpHeader.Controls.Add(this.pnlDeliveryPoint, 1, 1);
            this.tlpHeader.Controls.Add(this.pnlCreatedAt, 2, 1);
            this.tlpHeader.Controls.Add(this.pnlDiscount, 3, 1);
            this.tlpHeader.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tlpHeader.Location = new System.Drawing.Point(6, 24);
            this.tlpHeader.Margin = new System.Windows.Forms.Padding(0);
            this.tlpHeader.Name = "tlpHeader";
            this.tlpHeader.RowCount = 2;
            this.tlpHeader.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tlpHeader.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tlpHeader.Size = new System.Drawing.Size(1401, 127);
            this.tlpHeader.TabIndex = 0;
            // 
            // pnlCompany
            // 
            this.pnlCompany.Controls.Add(this.cmbCompany);
            this.pnlCompany.Controls.Add(this.lblCompany);
            this.pnlCompany.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlCompany.Location = new System.Drawing.Point(3, 3);
            this.pnlCompany.Name = "pnlCompany";
            this.pnlCompany.Padding = new System.Windows.Forms.Padding(6, 4, 6, 4);
            this.pnlCompany.Size = new System.Drawing.Size(301, 57);
            this.pnlCompany.TabIndex = 0;
            // 
            // cmbCompany
            // 
            this.cmbCompany.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.cmbCompany.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.cmbCompany.FormattingEnabled = true;
            this.cmbCompany.Location = new System.Drawing.Point(6, 25);
            this.cmbCompany.Name = "cmbCompany";
            this.cmbCompany.Size = new System.Drawing.Size(289, 28);
            this.cmbCompany.TabIndex = 1;
            this.cmbCompany.SelectedIndexChanged += new System.EventHandler(this.cmbCompany_SelectedIndexChanged);
            // 
            // lblCompany
            // 
            this.lblCompany.AutoSize = true;
            this.lblCompany.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblCompany.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblCompany.Location = new System.Drawing.Point(6, 4);
            this.lblCompany.Name = "lblCompany";
            this.lblCompany.Padding = new System.Windows.Forms.Padding(0, 0, 0, 2);
            this.lblCompany.Size = new System.Drawing.Size(75, 22);
            this.lblCompany.TabIndex = 2;
            this.lblCompany.Text = "Company:";
            // 
            // pnlClient
            // 
            this.pnlClient.Controls.Add(this.cmbClient);
            this.pnlClient.Controls.Add(this.lblClient);
            this.pnlClient.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlClient.Location = new System.Drawing.Point(310, 3);
            this.pnlClient.Name = "pnlClient";
            this.pnlClient.Padding = new System.Windows.Forms.Padding(6, 4, 6, 4);
            this.pnlClient.Size = new System.Drawing.Size(280, 57);
            this.pnlClient.TabIndex = 1;
            // 
            // cmbClient
            // 
            this.cmbClient.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.cmbClient.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.cmbClient.FormattingEnabled = true;
            this.cmbClient.Location = new System.Drawing.Point(6, 25);
            this.cmbClient.Name = "cmbClient";
            this.cmbClient.Size = new System.Drawing.Size(268, 28);
            this.cmbClient.TabIndex = 3;
            // 
            // lblClient
            // 
            this.lblClient.AutoSize = true;
            this.lblClient.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblClient.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblClient.Location = new System.Drawing.Point(6, 4);
            this.lblClient.Name = "lblClient";
            this.lblClient.Padding = new System.Windows.Forms.Padding(0, 0, 0, 2);
            this.lblClient.Size = new System.Drawing.Size(50, 22);
            this.lblClient.TabIndex = 4;
            this.lblClient.Text = "Client:";
            // 
            // pnlRFQCode
            // 
            this.pnlRFQCode.Controls.Add(this.txtRFQCode);
            this.pnlRFQCode.Controls.Add(this.lblRFQCode);
            this.pnlRFQCode.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlRFQCode.Location = new System.Drawing.Point(596, 3);
            this.pnlRFQCode.Name = "pnlRFQCode";
            this.pnlRFQCode.Padding = new System.Windows.Forms.Padding(6, 4, 6, 4);
            this.pnlRFQCode.Size = new System.Drawing.Size(295, 57);
            this.pnlRFQCode.TabIndex = 2;
            // 
            // txtRFQCode
            // 
            this.txtRFQCode.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.txtRFQCode.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtRFQCode.Location = new System.Drawing.Point(6, 26);
            this.txtRFQCode.Name = "txtRFQCode";
            this.txtRFQCode.Size = new System.Drawing.Size(283, 27);
            this.txtRFQCode.TabIndex = 5;
            // 
            // lblRFQCode
            // 
            this.lblRFQCode.AutoSize = true;
            this.lblRFQCode.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblRFQCode.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblRFQCode.Location = new System.Drawing.Point(6, 4);
            this.lblRFQCode.Name = "lblRFQCode";
            this.lblRFQCode.Padding = new System.Windows.Forms.Padding(0, 0, 0, 2);
            this.lblRFQCode.Size = new System.Drawing.Size(78, 22);
            this.lblRFQCode.TabIndex = 6;
            this.lblRFQCode.Text = "RFQ Code:";
            // 
            // pnlQuoteCode
            // 
            this.pnlQuoteCode.Controls.Add(this.txtQuoteCode);
            this.pnlQuoteCode.Controls.Add(this.lblQuoteCode);
            this.pnlQuoteCode.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlQuoteCode.Location = new System.Drawing.Point(897, 3);
            this.pnlQuoteCode.Name = "pnlQuoteCode";
            this.pnlQuoteCode.Padding = new System.Windows.Forms.Padding(6, 4, 6, 4);
            this.pnlQuoteCode.Size = new System.Drawing.Size(262, 57);
            this.pnlQuoteCode.TabIndex = 3;
            // 
            // txtQuoteCode
            // 
            this.txtQuoteCode.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.txtQuoteCode.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtQuoteCode.Location = new System.Drawing.Point(6, 26);
            this.txtQuoteCode.Name = "txtQuoteCode";
            this.txtQuoteCode.Size = new System.Drawing.Size(250, 27);
            this.txtQuoteCode.TabIndex = 7;
            // 
            // lblQuoteCode
            // 
            this.lblQuoteCode.AutoSize = true;
            this.lblQuoteCode.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblQuoteCode.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblQuoteCode.Location = new System.Drawing.Point(6, 4);
            this.lblQuoteCode.Name = "lblQuoteCode";
            this.lblQuoteCode.Padding = new System.Windows.Forms.Padding(0, 0, 0, 2);
            this.lblQuoteCode.Size = new System.Drawing.Size(92, 22);
            this.lblQuoteCode.TabIndex = 8;
            this.lblQuoteCode.Text = "Quote Code:";
            // 
            // pnlValidity
            // 
            this.pnlValidity.Controls.Add(this.txtValidity);
            this.pnlValidity.Controls.Add(this.lblValidity);
            this.pnlValidity.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlValidity.Location = new System.Drawing.Point(1165, 3);
            this.pnlValidity.Name = "pnlValidity";
            this.pnlValidity.Padding = new System.Windows.Forms.Padding(6, 4, 6, 4);
            this.pnlValidity.Size = new System.Drawing.Size(233, 57);
            this.pnlValidity.TabIndex = 4;
            // 
            // txtValidity
            // 
            this.txtValidity.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.txtValidity.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtValidity.Location = new System.Drawing.Point(6, 26);
            this.txtValidity.Name = "txtValidity";
            this.txtValidity.Size = new System.Drawing.Size(221, 27);
            this.txtValidity.TabIndex = 15;
            // 
            // lblValidity
            // 
            this.lblValidity.AutoSize = true;
            this.lblValidity.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblValidity.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblValidity.Location = new System.Drawing.Point(6, 4);
            this.lblValidity.Name = "lblValidity";
            this.lblValidity.Padding = new System.Windows.Forms.Padding(0, 0, 0, 2);
            this.lblValidity.Size = new System.Drawing.Size(107, 22);
            this.lblValidity.TabIndex = 16;
            this.lblValidity.Text = "Validity (Days):";
            // 
            // pnlDeliveryTerm
            // 
            this.pnlDeliveryTerm.Controls.Add(this.rbtnPCG);
            this.pnlDeliveryTerm.Controls.Add(this.rbtnDDUDAP);
            this.pnlDeliveryTerm.Controls.Add(this.rbtnDDP);
            this.pnlDeliveryTerm.Controls.Add(this.rbtnDAP);
            this.pnlDeliveryTerm.Controls.Add(this.lblDeliveryTerm);
            this.pnlDeliveryTerm.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlDeliveryTerm.Location = new System.Drawing.Point(3, 66);
            this.pnlDeliveryTerm.Name = "pnlDeliveryTerm";
            this.pnlDeliveryTerm.Padding = new System.Windows.Forms.Padding(6, 4, 6, 4);
            this.pnlDeliveryTerm.Size = new System.Drawing.Size(301, 58);
            this.pnlDeliveryTerm.TabIndex = 5;
            // 
            // rbtnPCG
            // 
            this.rbtnPCG.AutoSize = true;
            this.rbtnPCG.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.rbtnPCG.Location = new System.Drawing.Point(236, 27);
            this.rbtnPCG.Name = "rbtnPCG";
            this.rbtnPCG.Size = new System.Drawing.Size(57, 24);
            this.rbtnPCG.TabIndex = 23;
            this.rbtnPCG.Text = "PCG";
            this.rbtnPCG.UseVisualStyleBackColor = true;
            // 
            // rbtnDDUDAP
            // 
            this.rbtnDDUDAP.AutoSize = true;
            this.rbtnDDUDAP.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.rbtnDDUDAP.Location = new System.Drawing.Point(137, 27);
            this.rbtnDDUDAP.Name = "rbtnDDUDAP";
            this.rbtnDDUDAP.Size = new System.Drawing.Size(97, 24);
            this.rbtnDDUDAP.TabIndex = 22;
            this.rbtnDDUDAP.Text = "DDU/DAP";
            this.rbtnDDUDAP.UseVisualStyleBackColor = true;
            // 
            // rbtnDDP
            // 
            this.rbtnDDP.AutoSize = true;
            this.rbtnDDP.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.rbtnDDP.Location = new System.Drawing.Point(71, 28);
            this.rbtnDDP.Name = "rbtnDDP";
            this.rbtnDDP.Size = new System.Drawing.Size(60, 24);
            this.rbtnDDP.TabIndex = 21;
            this.rbtnDDP.Text = "DDP";
            this.rbtnDDP.UseVisualStyleBackColor = true;
            // 
            // rbtnDAP
            // 
            this.rbtnDAP.AutoSize = true;
            this.rbtnDAP.Checked = true;
            this.rbtnDAP.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.rbtnDAP.Location = new System.Drawing.Point(6, 28);
            this.rbtnDAP.Name = "rbtnDAP";
            this.rbtnDAP.Size = new System.Drawing.Size(59, 24);
            this.rbtnDAP.TabIndex = 20;
            this.rbtnDAP.TabStop = true;
            this.rbtnDAP.Text = "DAP";
            this.rbtnDAP.UseVisualStyleBackColor = true;
            // 
            // lblDeliveryTerm
            // 
            this.lblDeliveryTerm.AutoSize = true;
            this.lblDeliveryTerm.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblDeliveryTerm.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblDeliveryTerm.Location = new System.Drawing.Point(6, 4);
            this.lblDeliveryTerm.Name = "lblDeliveryTerm";
            this.lblDeliveryTerm.Padding = new System.Windows.Forms.Padding(0, 0, 0, 2);
            this.lblDeliveryTerm.Size = new System.Drawing.Size(103, 22);
            this.lblDeliveryTerm.TabIndex = 24;
            this.lblDeliveryTerm.Text = "Delivery Term:";
            // 
            // pnlDeliveryPoint
            // 
            this.pnlDeliveryPoint.Controls.Add(this.txtDeliveryPoint);
            this.pnlDeliveryPoint.Controls.Add(this.lblDeliveryPoint);
            this.pnlDeliveryPoint.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlDeliveryPoint.Location = new System.Drawing.Point(310, 66);
            this.pnlDeliveryPoint.Name = "pnlDeliveryPoint";
            this.pnlDeliveryPoint.Padding = new System.Windows.Forms.Padding(6, 4, 6, 4);
            this.pnlDeliveryPoint.Size = new System.Drawing.Size(280, 58);
            this.pnlDeliveryPoint.TabIndex = 6;
            // 
            // txtDeliveryPoint
            // 
            this.txtDeliveryPoint.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.txtDeliveryPoint.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtDeliveryPoint.Location = new System.Drawing.Point(6, 27);
            this.txtDeliveryPoint.Name = "txtDeliveryPoint";
            this.txtDeliveryPoint.Size = new System.Drawing.Size(268, 27);
            this.txtDeliveryPoint.TabIndex = 9;
            // 
            // lblDeliveryPoint
            // 
            this.lblDeliveryPoint.AutoSize = true;
            this.lblDeliveryPoint.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblDeliveryPoint.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblDeliveryPoint.Location = new System.Drawing.Point(6, 4);
            this.lblDeliveryPoint.Name = "lblDeliveryPoint";
            this.lblDeliveryPoint.Padding = new System.Windows.Forms.Padding(0, 0, 0, 2);
            this.lblDeliveryPoint.Size = new System.Drawing.Size(103, 22);
            this.lblDeliveryPoint.TabIndex = 10;
            this.lblDeliveryPoint.Text = "Delivery Point:";
            // 
            // pnlCreatedAt
            // 
            this.pnlCreatedAt.Controls.Add(this.dtpCreatedAt);
            this.pnlCreatedAt.Controls.Add(this.lblCreatedAt);
            this.pnlCreatedAt.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlCreatedAt.Location = new System.Drawing.Point(596, 66);
            this.pnlCreatedAt.Name = "pnlCreatedAt";
            this.pnlCreatedAt.Padding = new System.Windows.Forms.Padding(6, 4, 6, 4);
            this.pnlCreatedAt.Size = new System.Drawing.Size(295, 58);
            this.pnlCreatedAt.TabIndex = 7;
            // 
            // dtpCreatedAt
            // 
            this.dtpCreatedAt.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dtpCreatedAt.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.dtpCreatedAt.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpCreatedAt.Location = new System.Drawing.Point(6, 27);
            this.dtpCreatedAt.Name = "dtpCreatedAt";
            this.dtpCreatedAt.Size = new System.Drawing.Size(283, 27);
            this.dtpCreatedAt.TabIndex = 11;
            // 
            // lblCreatedAt
            // 
            this.lblCreatedAt.AutoSize = true;
            this.lblCreatedAt.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblCreatedAt.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblCreatedAt.Location = new System.Drawing.Point(6, 4);
            this.lblCreatedAt.Name = "lblCreatedAt";
            this.lblCreatedAt.Padding = new System.Windows.Forms.Padding(0, 0, 0, 2);
            this.lblCreatedAt.Size = new System.Drawing.Size(100, 22);
            this.lblCreatedAt.TabIndex = 12;
            this.lblCreatedAt.Text = "Created Date:";
            // 
            // pnlDiscount
            // 
            this.tlpHeader.SetColumnSpan(this.pnlDiscount, 2);
            this.pnlDiscount.Controls.Add(this.numDiscount);
            this.pnlDiscount.Controls.Add(this.lblDiscount);
            this.pnlDiscount.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlDiscount.Location = new System.Drawing.Point(897, 66);
            this.pnlDiscount.Name = "pnlDiscount";
            this.pnlDiscount.Padding = new System.Windows.Forms.Padding(6, 4, 6, 4);
            this.pnlDiscount.Size = new System.Drawing.Size(501, 58);
            this.pnlDiscount.TabIndex = 8;
            // 
            // numDiscount
            // 
            this.numDiscount.DecimalPlaces = 2;
            this.numDiscount.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.numDiscount.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.numDiscount.Location = new System.Drawing.Point(6, 27);
            this.numDiscount.Maximum = new decimal(new int[] {
            1000000,
            0,
            0,
            0});
            this.numDiscount.Name = "numDiscount";
            this.numDiscount.Size = new System.Drawing.Size(489, 27);
            this.numDiscount.TabIndex = 17;
            this.numDiscount.ValueChanged += new System.EventHandler(this.numDiscount_ValueChanged);
            // 
            // lblDiscount
            // 
            this.lblDiscount.AutoSize = true;
            this.lblDiscount.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblDiscount.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblDiscount.Location = new System.Drawing.Point(6, 4);
            this.lblDiscount.Name = "lblDiscount";
            this.lblDiscount.Padding = new System.Windows.Forms.Padding(0, 0, 0, 2);
            this.lblDiscount.Size = new System.Drawing.Size(96, 22);
            this.lblDiscount.TabIndex = 18;
            this.lblDiscount.Text = "Discount (%):";
            // 
            // grpCurrentItem
            // 
            this.grpCurrentItem.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.grpCurrentItem.Controls.Add(this.txtItemDescription);
            this.grpCurrentItem.Controls.Add(this.lblItemDescription);
            this.grpCurrentItem.Controls.Add(this.numItemQuantity);
            this.grpCurrentItem.Controls.Add(this.lblItemQuantity);
            this.grpCurrentItem.Controls.Add(this.txtItemUnit);
            this.grpCurrentItem.Controls.Add(this.lblItemUnit);
            this.grpCurrentItem.Controls.Add(this.numItemUnitPrice);
            this.grpCurrentItem.Controls.Add(this.lblItemUnitPrice);
            this.grpCurrentItem.Controls.Add(this.cmbCurrency);
            this.grpCurrentItem.Controls.Add(this.lblCurrency);
            this.grpCurrentItem.Controls.Add(this.numItemDeliveryTime);
            this.grpCurrentItem.Controls.Add(this.lblItemDeliveryTime);
            this.grpCurrentItem.Controls.Add(this.btnAddItem);
            this.grpCurrentItem.Controls.Add(this.btnClearItem);
            this.grpCurrentItem.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.grpCurrentItem.Location = new System.Drawing.Point(16, 178);
            this.grpCurrentItem.Name = "grpCurrentItem";
            this.grpCurrentItem.Padding = new System.Windows.Forms.Padding(4);
            this.grpCurrentItem.Size = new System.Drawing.Size(754, 380);
            this.grpCurrentItem.TabIndex = 1;
            this.grpCurrentItem.TabStop = false;
            this.grpCurrentItem.Text = "Item Details";
            // 
            // txtItemDescription
            // 
            this.txtItemDescription.AcceptsReturn = true;
            this.txtItemDescription.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtItemDescription.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtItemDescription.Location = new System.Drawing.Point(14, 44);
            this.txtItemDescription.Multiline = true;
            this.txtItemDescription.Name = "txtItemDescription";
            this.txtItemDescription.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtItemDescription.Size = new System.Drawing.Size(722, 200);
            this.txtItemDescription.TabIndex = 3;
            // 
            // lblItemDescription
            // 
            this.lblItemDescription.AutoSize = true;
            this.lblItemDescription.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblItemDescription.Location = new System.Drawing.Point(14, 22);
            this.lblItemDescription.Name = "lblItemDescription";
            this.lblItemDescription.Size = new System.Drawing.Size(88, 20);
            this.lblItemDescription.TabIndex = 2;
            this.lblItemDescription.Text = "Description:";
            // 
            // numItemQuantity
            // 
            this.numItemQuantity.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.numItemQuantity.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.numItemQuantity.Location = new System.Drawing.Point(14, 282);
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
            this.numItemQuantity.Size = new System.Drawing.Size(130, 27);
            this.numItemQuantity.TabIndex = 5;
            this.numItemQuantity.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // lblItemQuantity
            // 
            this.lblItemQuantity.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblItemQuantity.AutoSize = true;
            this.lblItemQuantity.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblItemQuantity.Location = new System.Drawing.Point(14, 258);
            this.lblItemQuantity.Name = "lblItemQuantity";
            this.lblItemQuantity.Size = new System.Drawing.Size(68, 20);
            this.lblItemQuantity.TabIndex = 4;
            this.lblItemQuantity.Text = "Quantity:";
            // 
            // txtItemUnit
            // 
            this.txtItemUnit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtItemUnit.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.txtItemUnit.Location = new System.Drawing.Point(160, 282);
            this.txtItemUnit.Name = "txtItemUnit";
            this.txtItemUnit.Size = new System.Drawing.Size(74, 27);
            this.txtItemUnit.TabIndex = 7;
            this.txtItemUnit.Text = "PCS";
            // 
            // lblItemUnit
            // 
            this.lblItemUnit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblItemUnit.AutoSize = true;
            this.lblItemUnit.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblItemUnit.Location = new System.Drawing.Point(160, 258);
            this.lblItemUnit.Name = "lblItemUnit";
            this.lblItemUnit.Size = new System.Drawing.Size(39, 20);
            this.lblItemUnit.TabIndex = 6;
            this.lblItemUnit.Text = "Unit:";
            // 
            // numItemUnitPrice
            // 
            this.numItemUnitPrice.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.numItemUnitPrice.DecimalPlaces = 2;
            this.numItemUnitPrice.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.numItemUnitPrice.Location = new System.Drawing.Point(250, 282);
            this.numItemUnitPrice.Maximum = new decimal(new int[] {
            1000000,
            0,
            0,
            0});
            this.numItemUnitPrice.Name = "numItemUnitPrice";
            this.numItemUnitPrice.Size = new System.Drawing.Size(174, 27);
            this.numItemUnitPrice.TabIndex = 9;
            // 
            // lblItemUnitPrice
            // 
            this.lblItemUnitPrice.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblItemUnitPrice.AutoSize = true;
            this.lblItemUnitPrice.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblItemUnitPrice.Location = new System.Drawing.Point(250, 258);
            this.lblItemUnitPrice.Name = "lblItemUnitPrice";
            this.lblItemUnitPrice.Size = new System.Drawing.Size(75, 20);
            this.lblItemUnitPrice.TabIndex = 8;
            this.lblItemUnitPrice.Text = "Unit Price:";
            // 
            // cmbCurrency
            // 
            this.cmbCurrency.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
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
            this.cmbCurrency.Location = new System.Drawing.Point(440, 282);
            this.cmbCurrency.Name = "cmbCurrency";
            this.cmbCurrency.Size = new System.Drawing.Size(90, 28);
            this.cmbCurrency.TabIndex = 19;
            // 
            // lblCurrency
            // 
            this.lblCurrency.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblCurrency.AutoSize = true;
            this.lblCurrency.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblCurrency.Location = new System.Drawing.Point(440, 258);
            this.lblCurrency.Name = "lblCurrency";
            this.lblCurrency.Size = new System.Drawing.Size(69, 20);
            this.lblCurrency.TabIndex = 18;
            this.lblCurrency.Text = "Currency:";
            // 
            // numItemDeliveryTime
            // 
            this.numItemDeliveryTime.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.numItemDeliveryTime.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.numItemDeliveryTime.Location = new System.Drawing.Point(546, 282);
            this.numItemDeliveryTime.Maximum = new decimal(new int[] {
            365,
            0,
            0,
            0});
            this.numItemDeliveryTime.Name = "numItemDeliveryTime";
            this.numItemDeliveryTime.Size = new System.Drawing.Size(184, 27);
            this.numItemDeliveryTime.TabIndex = 11;
            // 
            // lblItemDeliveryTime
            // 
            this.lblItemDeliveryTime.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblItemDeliveryTime.AutoSize = true;
            this.lblItemDeliveryTime.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lblItemDeliveryTime.Location = new System.Drawing.Point(546, 258);
            this.lblItemDeliveryTime.Name = "lblItemDeliveryTime";
            this.lblItemDeliveryTime.Size = new System.Drawing.Size(159, 20);
            this.lblItemDeliveryTime.TabIndex = 10;
            this.lblItemDeliveryTime.Text = "Delivery Time (Weeks):";
            // 
            // btnAddItem
            // 
            this.btnAddItem.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnAddItem.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnAddItem.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnAddItem.Location = new System.Drawing.Point(14, 328);
            this.btnAddItem.Name = "btnAddItem";
            this.btnAddItem.Size = new System.Drawing.Size(160, 37);
            this.btnAddItem.TabIndex = 0;
            this.btnAddItem.Text = "Add Item";
            this.btnAddItem.UseVisualStyleBackColor = false;
            this.btnAddItem.Click += new System.EventHandler(this.btnAddItem_Click);
            // 
            // btnClearItem
            // 
            this.btnClearItem.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnClearItem.BackColor = System.Drawing.Color.DarkSalmon;
            this.btnClearItem.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnClearItem.Location = new System.Drawing.Point(186, 328);
            this.btnClearItem.Name = "btnClearItem";
            this.btnClearItem.Size = new System.Drawing.Size(160, 37);
            this.btnClearItem.TabIndex = 1;
            this.btnClearItem.Text = "Clear";
            this.btnClearItem.UseVisualStyleBackColor = false;
            this.btnClearItem.Click += new System.EventHandler(this.btnClearItem_Click);
            // 
            // grpItemsList
            // 
            this.grpItemsList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grpItemsList.Controls.Add(this.lstItems);
            this.grpItemsList.Controls.Add(this.btnEditItem);
            this.grpItemsList.Controls.Add(this.btnRemoveItem);
            this.grpItemsList.Controls.Add(this.lblItemCount);
            this.grpItemsList.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.grpItemsList.Location = new System.Drawing.Point(785, 178);
            this.grpItemsList.Name = "grpItemsList";
            this.grpItemsList.Padding = new System.Windows.Forms.Padding(4);
            this.grpItemsList.Size = new System.Drawing.Size(645, 380);
            this.grpItemsList.TabIndex = 2;
            this.grpItemsList.TabStop = false;
            this.grpItemsList.Text = "Items List";
            // 
            // lstItems
            // 
            this.lstItems.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lstItems.Font = new System.Drawing.Font("Consolas", 8.25F);
            this.lstItems.FormattingEnabled = true;
            this.lstItems.ItemHeight = 17;
            this.lstItems.Location = new System.Drawing.Point(14, 26);
            this.lstItems.Name = "lstItems";
            this.lstItems.Size = new System.Drawing.Size(613, 259);
            this.lstItems.TabIndex = 0;
            this.lstItems.SelectedIndexChanged += new System.EventHandler(this.lstItems_SelectedIndexChanged);
            // 
            // btnEditItem
            // 
            this.btnEditItem.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnEditItem.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnEditItem.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnEditItem.Location = new System.Drawing.Point(14, 330);
            this.btnEditItem.Name = "btnEditItem";
            this.btnEditItem.Size = new System.Drawing.Size(133, 37);
            this.btnEditItem.TabIndex = 1;
            this.btnEditItem.Text = "Edit";
            this.btnEditItem.UseVisualStyleBackColor = false;
            this.btnEditItem.Click += new System.EventHandler(this.btnEditItem_Click);
            // 
            // btnRemoveItem
            // 
            this.btnRemoveItem.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnRemoveItem.BackColor = System.Drawing.Color.DarkSalmon;
            this.btnRemoveItem.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnRemoveItem.Location = new System.Drawing.Point(157, 330);
            this.btnRemoveItem.Name = "btnRemoveItem";
            this.btnRemoveItem.Size = new System.Drawing.Size(133, 37);
            this.btnRemoveItem.TabIndex = 2;
            this.btnRemoveItem.Text = "Remove";
            this.btnRemoveItem.UseVisualStyleBackColor = false;
            this.btnRemoveItem.Click += new System.EventHandler(this.btnRemoveItem_Click);
            // 
            // lblItemCount
            // 
            this.lblItemCount.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblItemCount.AutoSize = true;
            this.lblItemCount.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblItemCount.ForeColor = System.Drawing.Color.Gray;
            this.lblItemCount.Location = new System.Drawing.Point(310, 338);
            this.lblItemCount.Name = "lblItemCount";
            this.lblItemCount.Size = new System.Drawing.Size(91, 19);
            this.lblItemCount.TabIndex = 3;
            this.lblItemCount.Text = "Total Items: 0";
            // 
            // grpActions
            // 
            this.grpActions.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grpActions.Controls.Add(this.btnNew);
            this.grpActions.Controls.Add(this.btnGenerateBothExcel);
            this.grpActions.Controls.Add(this.btnGenerateBothPDF);
            this.grpActions.Controls.Add(this.lblTemplateLabel);
            this.grpActions.Controls.Add(this.lblTemplatePath);
            this.grpActions.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.grpActions.Location = new System.Drawing.Point(16, 566);
            this.grpActions.Name = "grpActions";
            this.grpActions.Padding = new System.Windows.Forms.Padding(4);
            this.grpActions.Size = new System.Drawing.Size(1413, 150);
            this.grpActions.TabIndex = 3;
            this.grpActions.TabStop = false;
            this.grpActions.Text = "Actions";
            // 
            // btnNew
            // 
            this.btnNew.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnNew.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnNew.Location = new System.Drawing.Point(14, 31);
            this.btnNew.Name = "btnNew";
            this.btnNew.Size = new System.Drawing.Size(133, 37);
            this.btnNew.TabIndex = 0;
            this.btnNew.Text = "New RFQ";
            this.btnNew.UseVisualStyleBackColor = false;
            this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // btnGenerateBothExcel
            // 
            this.btnGenerateBothExcel.BackColor = System.Drawing.Color.LightGreen;
            this.btnGenerateBothExcel.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnGenerateBothExcel.ForeColor = System.Drawing.Color.Black;
            this.btnGenerateBothExcel.Location = new System.Drawing.Point(157, 31);
            this.btnGenerateBothExcel.Name = "btnGenerateBothExcel";
            this.btnGenerateBothExcel.Size = new System.Drawing.Size(200, 37);
            this.btnGenerateBothExcel.TabIndex = 9;
            this.btnGenerateBothExcel.Text = "Generate Excel";
            this.btnGenerateBothExcel.UseVisualStyleBackColor = false;
            this.btnGenerateBothExcel.Click += new System.EventHandler(this.btnGenerateBothExcel_Click);
            // 
            // btnGenerateBothPDF
            // 
            this.btnGenerateBothPDF.BackColor = System.Drawing.Color.MediumPurple;
            this.btnGenerateBothPDF.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.btnGenerateBothPDF.ForeColor = System.Drawing.Color.White;
            this.btnGenerateBothPDF.Location = new System.Drawing.Point(367, 31);
            this.btnGenerateBothPDF.Name = "btnGenerateBothPDF";
            this.btnGenerateBothPDF.Size = new System.Drawing.Size(200, 37);
            this.btnGenerateBothPDF.TabIndex = 10;
            this.btnGenerateBothPDF.Text = "Generate PDF";
            this.btnGenerateBothPDF.UseVisualStyleBackColor = false;
            this.btnGenerateBothPDF.Click += new System.EventHandler(this.btnGenerateBothPDF_Click);
            // 
            // lblTemplateLabel
            // 
            this.lblTemplateLabel.AutoSize = true;
            this.lblTemplateLabel.Font = new System.Drawing.Font("Segoe UI", 8F);
            this.lblTemplateLabel.Location = new System.Drawing.Point(14, 118);
            this.lblTemplateLabel.Name = "lblTemplateLabel";
            this.lblTemplateLabel.Size = new System.Drawing.Size(67, 19);
            this.lblTemplateLabel.TabIndex = 3;
            this.lblTemplateLabel.Text = "Template:";
            // 
            // lblTemplatePath
            // 
            this.lblTemplatePath.AutoSize = true;
            this.lblTemplatePath.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Italic);
            this.lblTemplatePath.ForeColor = System.Drawing.Color.Gray;
            this.lblTemplatePath.Location = new System.Drawing.Point(87, 118);
            this.lblTemplatePath.Name = "lblTemplatePath";
            this.lblTemplatePath.Size = new System.Drawing.Size(142, 19);
            this.lblTemplatePath.TabIndex = 4;
            this.lblTemplatePath.Text = "No template selected";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1445, 732);
            this.Controls.Add(this.grpActions);
            this.Controls.Add(this.grpItemsList);
            this.Controls.Add(this.grpCurrentItem);
            this.Controls.Add(this.grpHeader);
            this.MinimumSize = new System.Drawing.Size(960, 700);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "RFQ Generator System";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.grpHeader.ResumeLayout(false);
            this.tlpHeader.ResumeLayout(false);
            this.pnlCompany.ResumeLayout(false);
            this.pnlCompany.PerformLayout();
            this.pnlClient.ResumeLayout(false);
            this.pnlClient.PerformLayout();
            this.pnlRFQCode.ResumeLayout(false);
            this.pnlRFQCode.PerformLayout();
            this.pnlQuoteCode.ResumeLayout(false);
            this.pnlQuoteCode.PerformLayout();
            this.pnlValidity.ResumeLayout(false);
            this.pnlValidity.PerformLayout();
            this.pnlDeliveryTerm.ResumeLayout(false);
            this.pnlDeliveryTerm.PerformLayout();
            this.pnlDeliveryPoint.ResumeLayout(false);
            this.pnlDeliveryPoint.PerformLayout();
            this.pnlCreatedAt.ResumeLayout(false);
            this.pnlCreatedAt.PerformLayout();
            this.pnlDiscount.ResumeLayout(false);
            this.pnlDiscount.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numDiscount)).EndInit();
            this.grpCurrentItem.ResumeLayout(false);
            this.grpCurrentItem.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numItemQuantity)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numItemUnitPrice)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numItemDeliveryTime)).EndInit();
            this.grpItemsList.ResumeLayout(false);
            this.grpItemsList.PerformLayout();
            this.grpActions.ResumeLayout(false);
            this.grpActions.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox grpHeader;
        private System.Windows.Forms.TableLayoutPanel tlpHeader;
        private System.Windows.Forms.Panel pnlCompany;
        private System.Windows.Forms.Panel pnlClient;
        private System.Windows.Forms.Panel pnlRFQCode;
        private System.Windows.Forms.Panel pnlQuoteCode;
        private System.Windows.Forms.Panel pnlValidity;
        private System.Windows.Forms.Panel pnlDeliveryTerm;
        private System.Windows.Forms.Panel pnlDeliveryPoint;
        private System.Windows.Forms.Panel pnlCreatedAt;
        private System.Windows.Forms.Panel pnlDiscount;
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
        private System.Windows.Forms.RadioButton rbtnDDUDAP;
        private System.Windows.Forms.RadioButton rbtnPCG;
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
        private System.Windows.Forms.Button btnGenerateBothExcel;
        private System.Windows.Forms.Button btnGenerateBothPDF;
        private System.Windows.Forms.Label lblTemplatePath;
        private System.Windows.Forms.Label lblTemplateLabel;
    }
}