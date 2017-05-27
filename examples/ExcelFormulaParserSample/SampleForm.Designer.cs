namespace ExcelFormulaParserSample
{
	partial class SampleForm
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
      System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
      System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
      System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
      System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
      this.ExcelFormulaTextBox = new System.Windows.Forms.TextBox();
      this.label1 = new System.Windows.Forms.Label();
      this.ParseButton = new System.Windows.Forms.Button();
      this.panel1 = new System.Windows.Forms.Panel();
      this.ExcelFormulaTokensGridView = new System.Windows.Forms.DataGridView();
      this.index = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.type = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.subtype = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.token = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.tokentree = new System.Windows.Forms.DataGridViewTextBoxColumn();
      this.panel1.SuspendLayout();
      ((System.ComponentModel.ISupportInitialize) (this.ExcelFormulaTokensGridView)).BeginInit();
      this.SuspendLayout();
      // 
      // ExcelFormulaTextBox
      // 
      this.ExcelFormulaTextBox.Location = new System.Drawing.Point(11, 33);
      this.ExcelFormulaTextBox.Name = "ExcelFormulaTextBox";
      this.ExcelFormulaTextBox.Size = new System.Drawing.Size(658, 21);
      this.ExcelFormulaTextBox.TabIndex = 0;
      this.ExcelFormulaTextBox.Enter += new System.EventHandler(this.ExcelFormulaTextBox_Enter);
      // 
      // label1
      // 
      this.label1.AutoSize = true;
      this.label1.Location = new System.Drawing.Point(11, 16);
      this.label1.Margin = new System.Windows.Forms.Padding(0);
      this.label1.Name = "label1";
      this.label1.Size = new System.Drawing.Size(75, 13);
      this.label1.TabIndex = 1;
      this.label1.Text = "Excel formula:";
      // 
      // ParseButton
      // 
      this.ParseButton.Location = new System.Drawing.Point(675, 31);
      this.ParseButton.Name = "ParseButton";
      this.ParseButton.Size = new System.Drawing.Size(65, 24);
      this.ParseButton.TabIndex = 2;
      this.ParseButton.Text = "Parse";
      this.ParseButton.UseVisualStyleBackColor = true;
      this.ParseButton.Click += new System.EventHandler(this.ParseButton_Click);
      // 
      // panel1
      // 
      this.panel1.AutoSize = true;
      this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
      this.panel1.Controls.Add(this.ExcelFormulaTokensGridView);
      this.panel1.Location = new System.Drawing.Point(11, 64);
      this.panel1.Name = "panel1";
      this.panel1.Size = new System.Drawing.Size(729, 580);
      this.panel1.TabIndex = 4;
      // 
      // ExcelFormulaTokensGridView
      // 
      this.ExcelFormulaTokensGridView.AllowUserToAddRows = false;
      this.ExcelFormulaTokensGridView.AllowUserToDeleteRows = false;
      this.ExcelFormulaTokensGridView.AllowUserToResizeRows = false;
      this.ExcelFormulaTokensGridView.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
      this.ExcelFormulaTokensGridView.BorderStyle = System.Windows.Forms.BorderStyle.None;
      this.ExcelFormulaTokensGridView.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.SingleHorizontal;
      this.ExcelFormulaTokensGridView.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
      dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
      dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
      dataGridViewCellStyle1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
      dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
      dataGridViewCellStyle1.Padding = new System.Windows.Forms.Padding(0, 2, 0, 2);
      dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
      dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
      dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
      this.ExcelFormulaTokensGridView.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
      this.ExcelFormulaTokensGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
      this.ExcelFormulaTokensGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.index,
            this.type,
            this.subtype,
            this.token,
            this.tokentree});
      dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
      dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
      dataGridViewCellStyle2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
      dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
      dataGridViewCellStyle2.Padding = new System.Windows.Forms.Padding(1, 0, 0, 0);
      dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Window;
      dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.ControlText;
      dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
      this.ExcelFormulaTokensGridView.DefaultCellStyle = dataGridViewCellStyle2;
      this.ExcelFormulaTokensGridView.Dock = System.Windows.Forms.DockStyle.Fill;
      this.ExcelFormulaTokensGridView.Location = new System.Drawing.Point(0, 0);
      this.ExcelFormulaTokensGridView.Name = "ExcelFormulaTokensGridView";
      this.ExcelFormulaTokensGridView.ReadOnly = true;
      dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
      dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
      dataGridViewCellStyle3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
      dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
      dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
      dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
      dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
      this.ExcelFormulaTokensGridView.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
      this.ExcelFormulaTokensGridView.RowHeadersVisible = false;
      dataGridViewCellStyle4.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
      this.ExcelFormulaTokensGridView.RowsDefaultCellStyle = dataGridViewCellStyle4;
      this.ExcelFormulaTokensGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
      this.ExcelFormulaTokensGridView.ShowCellErrors = false;
      this.ExcelFormulaTokensGridView.ShowCellToolTips = false;
      this.ExcelFormulaTokensGridView.ShowEditingIcon = false;
      this.ExcelFormulaTokensGridView.ShowRowErrors = false;
      this.ExcelFormulaTokensGridView.Size = new System.Drawing.Size(725, 576);
      this.ExcelFormulaTokensGridView.StandardTab = true;
      this.ExcelFormulaTokensGridView.TabIndex = 4;
      // 
      // index
      // 
      this.index.HeaderText = "index";
      this.index.Name = "index";
      this.index.ReadOnly = true;
      this.index.Width = 50;
      // 
      // type
      // 
      this.type.DataPropertyName = "Type";
      this.type.HeaderText = "type";
      this.type.Name = "type";
      this.type.ReadOnly = true;
      this.type.Width = 130;
      // 
      // subtype
      // 
      this.subtype.DataPropertyName = "Subtype";
      this.subtype.HeaderText = "subtype";
      this.subtype.Name = "subtype";
      this.subtype.ReadOnly = true;
      this.subtype.Width = 120;
      // 
      // token
      // 
      this.token.DataPropertyName = "Value";
      this.token.HeaderText = "token";
      this.token.Name = "token";
      this.token.ReadOnly = true;
      this.token.Width = 200;
      // 
      // tokentree
      // 
      this.tokentree.HeaderText = "tokentree";
      this.tokentree.Name = "tokentree";
      this.tokentree.ReadOnly = true;
      this.tokentree.Width = 225;
      // 
      // SampleForm
      // 
      this.AcceptButton = this.ParseButton;
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(750, 656);
      this.Controls.Add(this.panel1);
      this.Controls.Add(this.ParseButton);
      this.Controls.Add(this.label1);
      this.Controls.Add(this.ExcelFormulaTextBox);
      this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
      this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
      this.MaximizeBox = false;
      this.Name = "SampleForm";
      this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
      this.Text = "ExcelFormulaParserSample";
      this.panel1.ResumeLayout(false);
      ((System.ComponentModel.ISupportInitialize) (this.ExcelFormulaTokensGridView)).EndInit();
      this.ResumeLayout(false);
      this.PerformLayout();

		}

		#endregion

    private System.Windows.Forms.TextBox ExcelFormulaTextBox;
    private System.Windows.Forms.Label label1;
    private System.Windows.Forms.Button ParseButton;
    private System.Windows.Forms.Panel panel1;
    private System.Windows.Forms.DataGridView ExcelFormulaTokensGridView;
    private System.Windows.Forms.DataGridViewTextBoxColumn index;
    private System.Windows.Forms.DataGridViewTextBoxColumn type;
    private System.Windows.Forms.DataGridViewTextBoxColumn subtype;
    private System.Windows.Forms.DataGridViewTextBoxColumn token;
    private System.Windows.Forms.DataGridViewTextBoxColumn tokentree;
	}
}

