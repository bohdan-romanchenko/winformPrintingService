namespace Zebra
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.idDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.namesDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.sizesDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.manufacturersDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.compositionsDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.additionalsDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.goodsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.dbDataSet = new Zebra.dbDataSet();
            this.goodsTableAdapter = new Zebra.dbDataSetTableAdapters.goodsTableAdapter();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.goodsBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dbDataSet)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.idDataGridViewTextBoxColumn,
            this.namesDataGridViewTextBoxColumn,
            this.sizesDataGridViewTextBoxColumn,
            this.manufacturersDataGridViewTextBoxColumn,
            this.compositionsDataGridViewTextBoxColumn,
            this.additionalsDataGridViewTextBoxColumn});
            this.dataGridView1.DataSource = this.goodsBindingSource;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.DefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.Location = new System.Drawing.Point(0, -1);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(644, 337);
            this.dataGridView1.TabIndex = 1;
            // 
            // idDataGridViewTextBoxColumn
            // 
            this.idDataGridViewTextBoxColumn.DataPropertyName = "id";
            this.idDataGridViewTextBoxColumn.HeaderText = "id";
            this.idDataGridViewTextBoxColumn.Name = "idDataGridViewTextBoxColumn";
            // 
            // namesDataGridViewTextBoxColumn
            // 
            this.namesDataGridViewTextBoxColumn.DataPropertyName = "Names";
            this.namesDataGridViewTextBoxColumn.HeaderText = "Names";
            this.namesDataGridViewTextBoxColumn.Name = "namesDataGridViewTextBoxColumn";
            // 
            // sizesDataGridViewTextBoxColumn
            // 
            this.sizesDataGridViewTextBoxColumn.DataPropertyName = "Sizes";
            this.sizesDataGridViewTextBoxColumn.HeaderText = "Sizes";
            this.sizesDataGridViewTextBoxColumn.Name = "sizesDataGridViewTextBoxColumn";
            // 
            // manufacturersDataGridViewTextBoxColumn
            // 
            this.manufacturersDataGridViewTextBoxColumn.DataPropertyName = "Manufacturers";
            this.manufacturersDataGridViewTextBoxColumn.HeaderText = "Manufacturers";
            this.manufacturersDataGridViewTextBoxColumn.Name = "manufacturersDataGridViewTextBoxColumn";
            // 
            // compositionsDataGridViewTextBoxColumn
            // 
            this.compositionsDataGridViewTextBoxColumn.DataPropertyName = "Compositions";
            this.compositionsDataGridViewTextBoxColumn.HeaderText = "Compositions";
            this.compositionsDataGridViewTextBoxColumn.Name = "compositionsDataGridViewTextBoxColumn";
            // 
            // additionalsDataGridViewTextBoxColumn
            // 
            this.additionalsDataGridViewTextBoxColumn.DataPropertyName = "Additionals";
            this.additionalsDataGridViewTextBoxColumn.HeaderText = "Additionals";
            this.additionalsDataGridViewTextBoxColumn.Name = "additionalsDataGridViewTextBoxColumn";
            // 
            // goodsBindingSource
            // 
            this.goodsBindingSource.DataMember = "goods";
            this.goodsBindingSource.DataSource = this.dbDataSet;
            // 
            // dbDataSet
            // 
            this.dbDataSet.DataSetName = "dbDataSet";
            this.dbDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // goodsTableAdapter
            // 
            this.goodsTableAdapter.ClearBeforeFill = true;
            // 
            // button3
            // 
            this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button3.Location = new System.Drawing.Point(329, 371);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(315, 32);
            this.button3.TabIndex = 4;
            this.button3.Text = "Зберегти зміни";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button2
            // 
            this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button2.Location = new System.Drawing.Point(0, 403);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(315, 32);
            this.button2.TabIndex = 6;
            this.button2.Text = "Друк";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(329, 403);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(315, 34);
            this.button1.TabIndex = 7;
            this.button1.Text = "Додати товари";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // button4
            // 
            this.button4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button4.Location = new System.Drawing.Point(0, 371);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(315, 32);
            this.button4.TabIndex = 8;
            this.button4.Text = "Надрукувати обрану позицію";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(329, 342);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(315, 32);
            this.button5.TabIndex = 9;
            this.button5.Text = "Знайти";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(644, 435);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Form1";
            this.Text = "Стартове вікно";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.goodsBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dbDataSet)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridViewTextBoxColumn назваТоваруDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn розмірЕтикеткиDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn виробникDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn складDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn додатковоDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridView dataGridView1;
        private dbDataSet dbDataSet;
        private System.Windows.Forms.BindingSource goodsBindingSource;
        private dbDataSetTableAdapters.goodsTableAdapter goodsTableAdapter;
        private System.Windows.Forms.DataGridViewTextBoxColumn idDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn namesDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn sizesDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn manufacturersDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn compositionsDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn additionalsDataGridViewTextBoxColumn;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;

    }
}

