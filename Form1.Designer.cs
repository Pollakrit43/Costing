namespace Costing
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.dgvColumns = new System.Windows.Forms.DataGridView();
            this.btnStart = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.cboStyle = new System.Windows.Forms.ComboBox();
            this.cboCustomer = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnReset = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.tb2 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cboType = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.Type = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Style = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Customer = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.F1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.F2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.F3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.F4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.F5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.F6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.F7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.F8 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.F9 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.F10 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.F11 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.F12 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgvColumns)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvColumns
            // 
            this.dgvColumns.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvColumns.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvColumns.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Type,
            this.Style,
            this.Customer,
            this.F1,
            this.F2,
            this.F3,
            this.F4,
            this.F5,
            this.F6,
            this.F7,
            this.F8,
            this.F9,
            this.F10,
            this.F11,
            this.F12});
            this.dgvColumns.Location = new System.Drawing.Point(12, 97);
            this.dgvColumns.Name = "dgvColumns";
            this.dgvColumns.Size = new System.Drawing.Size(776, 267);
            this.dgvColumns.TabIndex = 0;
            this.dgvColumns.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvColumns_CellClick);
            this.dgvColumns.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvColumns_CellEndEdit);
            this.dgvColumns.CellMouseDown += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgvColumns_CellMouseDown);
            this.dgvColumns.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvColumns_CellValueChanged);
            this.dgvColumns.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.dgvColumns_RowPostPaint);
            // 
            // btnStart
            // 
            this.btnStart.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnStart.Location = new System.Drawing.Point(492, 27);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(75, 23);
            this.btnStart.TabIndex = 1;
            this.btnStart.Text = "Start";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // btnSave
            // 
            this.btnSave.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSave.Location = new System.Drawing.Point(690, 27);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 2;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // cboStyle
            // 
            this.cboStyle.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cboStyle.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.cboStyle.FormattingEnabled = true;
            this.cboStyle.Location = new System.Drawing.Point(141, 29);
            this.cboStyle.Name = "cboStyle";
            this.cboStyle.Size = new System.Drawing.Size(121, 21);
            this.cboStyle.TabIndex = 3;
            this.cboStyle.SelectedIndexChanged += new System.EventHandler(this.cboStyle_SelectedIndexChanged);
            // 
            // cboCustomer
            // 
            this.cboCustomer.FormattingEnabled = true;
            this.cboCustomer.Location = new System.Drawing.Point(281, 29);
            this.cboCustomer.Name = "cboCustomer";
            this.cboCustomer.Size = new System.Drawing.Size(121, 21);
            this.cboCustomer.TabIndex = 4;
            this.cboCustomer.SelectedIndexChanged += new System.EventHandler(this.cboCustomer_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(139, 5);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(49, 20);
            this.label1.TabIndex = 5;
            this.label1.Text = "Style";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(277, 4);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(86, 20);
            this.label2.TabIndex = 6;
            this.label2.Text = "Customer";
            // 
            // btnReset
            // 
            this.btnReset.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnReset.Location = new System.Drawing.Point(573, 27);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(75, 23);
            this.btnReset.TabIndex = 7;
            this.btnReset.Text = "Reset";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(12, 71);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(443, 20);
            this.textBox1.TabIndex = 8;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // tb2
            // 
            this.tb2.Location = new System.Drawing.Point(408, 29);
            this.tb2.Name = "tb2";
            this.tb2.Size = new System.Drawing.Size(56, 20);
            this.tb2.TabIndex = 9;
            this.tb2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.tb2_KeyPress);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(405, 4);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(24, 20);
            this.label3.TabIndex = 10;
            this.label3.Text = "%";
            // 
            // cboType
            // 
            this.cboType.FormattingEnabled = true;
            this.cboType.Location = new System.Drawing.Point(12, 29);
            this.cboType.Name = "cboType";
            this.cboType.Size = new System.Drawing.Size(121, 21);
            this.cboType.TabIndex = 11;
            this.cboType.SelectedIndexChanged += new System.EventHandler(this.cboType_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(10, 6);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(47, 20);
            this.label4.TabIndex = 12;
            this.label4.Text = "Type";
            // 
            // Type
            // 
            this.Type.HeaderText = "Type";
            this.Type.Name = "Type";
            this.Type.Visible = false;
            // 
            // Style
            // 
            this.Style.HeaderText = "Style";
            this.Style.Name = "Style";
            this.Style.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Style.Visible = false;
            // 
            // Customer
            // 
            this.Customer.HeaderText = "Customer";
            this.Customer.Name = "Customer";
            this.Customer.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.Customer.Visible = false;
            // 
            // F1
            // 
            this.F1.HeaderText = "F1";
            this.F1.Name = "F1";
            this.F1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // F2
            // 
            this.F2.HeaderText = "F2";
            this.F2.Name = "F2";
            this.F2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // F3
            // 
            this.F3.HeaderText = "F3";
            this.F3.Name = "F3";
            this.F3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // F4
            // 
            this.F4.HeaderText = "F4";
            this.F4.Name = "F4";
            this.F4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // F5
            // 
            this.F5.HeaderText = "F5";
            this.F5.Name = "F5";
            this.F5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // F6
            // 
            this.F6.HeaderText = "F6";
            this.F6.Name = "F6";
            this.F6.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // F7
            // 
            this.F7.HeaderText = "F7";
            this.F7.Name = "F7";
            this.F7.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // F8
            // 
            this.F8.HeaderText = "F8";
            this.F8.Name = "F8";
            this.F8.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // F9
            // 
            this.F9.HeaderText = "F9";
            this.F9.Name = "F9";
            this.F9.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // F10
            // 
            this.F10.HeaderText = "F10";
            this.F10.Name = "F10";
            this.F10.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // F11
            // 
            this.F11.HeaderText = "F11";
            this.F11.Name = "F11";
            this.F11.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // F12
            // 
            this.F12.HeaderText = "F12";
            this.F12.Name = "F12";
            this.F12.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.RoyalBlue;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cboType);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.tb2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cboCustomer);
            this.Controls.Add(this.cboStyle);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.dgvColumns);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Costing";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvColumns)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvColumns;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.ComboBox cboStyle;
        private System.Windows.Forms.ComboBox cboCustomer;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox tb2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cboType;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Type;
        private System.Windows.Forms.DataGridViewTextBoxColumn Style;
        private System.Windows.Forms.DataGridViewTextBoxColumn Customer;
        private System.Windows.Forms.DataGridViewTextBoxColumn F1;
        private System.Windows.Forms.DataGridViewTextBoxColumn F2;
        private System.Windows.Forms.DataGridViewTextBoxColumn F3;
        private System.Windows.Forms.DataGridViewTextBoxColumn F4;
        private System.Windows.Forms.DataGridViewTextBoxColumn F5;
        private System.Windows.Forms.DataGridViewTextBoxColumn F6;
        private System.Windows.Forms.DataGridViewTextBoxColumn F7;
        private System.Windows.Forms.DataGridViewTextBoxColumn F8;
        private System.Windows.Forms.DataGridViewTextBoxColumn F9;
        private System.Windows.Forms.DataGridViewTextBoxColumn F10;
        private System.Windows.Forms.DataGridViewTextBoxColumn F11;
        private System.Windows.Forms.DataGridViewTextBoxColumn F12;
    }
}

